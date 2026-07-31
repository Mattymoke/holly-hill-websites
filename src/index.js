// Holly Hill Surplus — shop backend (Cloudflare Worker)
//
// Routes:
//   GET  /api/lots              -> list available lots (public)
//   POST /api/checkout          -> create a Stripe Checkout session (requires login)
//   POST /api/webhook/stripe    -> Stripe calls this when payment completes
//   GET  /api/orders            -> order history for the logged-in buyer
//
// Anything that isn't one of these paths falls through to the static
// site (index.html, /mushrooms, /surplus) automatically -- that's what
// the "assets" config in wrangler.jsonc does, no code needed here for it.
//
// Requires these secrets to be set (see the setup notes at the bottom):
//   wrangler secret put STRIPE_SECRET_KEY
//   wrangler secret put STRIPE_WEBHOOK_SECRET
//   wrangler secret put CLERK_SECRET_KEY

import { createClerkClient } from "@clerk/backend";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json" },
  });
}

async function getAuthedUserId(request, env) {
  const clerk = createClerkClient({ secretKey: env.CLERK_SECRET_KEY });
  try {
    const { isSignedIn, toAuth } = await clerk.authenticateRequest(request, {
      // Adjust to your real domain once this is live
      authorizedParties: ["https://hollyhillohio.com"],
    });
    if (!isSignedIn) return null;
    return toAuth().userId;
  } catch (err) {
    return null;
  }
}

async function handleListLots(env) {
  const { results } = await env.DB.prepare(
    "SELECT id, name, category, description, website_price_cents FROM lots WHERE status = 'available' ORDER BY created_at DESC"
  ).all();
  return json({ lots: results });
}

async function handleCheckout(request, env) {
  const userId = await getAuthedUserId(request, env);
  if (!userId) return json({ error: "Please log in first" }, 401);

  const { lotId } = await request.json();
  if (!lotId) return json({ error: "lotId is required" }, 400);

  const lot = await env.DB.prepare(
    "SELECT * FROM lots WHERE id = ? AND status = 'available'"
  )
    .bind(lotId)
    .first();

  if (!lot) {
    return json({ error: "That lot is no longer available" }, 409);
  }

  // Mark it "reserved" immediately so a second buyer can't also check out
  // while this buyer is on the Stripe payment page.
  await env.DB.prepare("UPDATE lots SET status = 'reserved', updated_at = datetime('now') WHERE id = ?")
    .bind(lotId)
    .run();

  const orderId = crypto.randomUUID();

  const params = new URLSearchParams({
    "mode": "payment",
    "success_url": "https://hollyhillohio.com/surplus/shop/success?order=" + orderId,
    "cancel_url": "https://hollyhillohio.com/surplus/shop",
    "line_items[0][quantity]": "1",
    "line_items[0][price_data][currency]": "usd",
    "line_items[0][price_data][unit_amount]": String(lot.website_price_cents),
    "line_items[0][price_data][product_data][name]": lot.name,
    "metadata[order_id]": orderId,
    "metadata[lot_id]": lotId,
    "metadata[clerk_user_id]": userId,
  });

  const stripeRes = await fetch("https://api.stripe.com/v1/checkout/sessions", {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + env.STRIPE_SECRET_KEY,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params,
  });

  if (!stripeRes.ok) {
    // Release the reservation if Stripe failed to create the session
    await env.DB.prepare("UPDATE lots SET status = 'available', updated_at = datetime('now') WHERE id = ?")
      .bind(lotId)
      .run();
    const errText = await stripeRes.text();
    return json({ error: "Could not start checkout", detail: errText }, 502);
  }

  const session = await stripeRes.json();

  await env.DB.prepare(
    `INSERT INTO orders (id, clerk_user_id, lot_id, amount_cents, status, stripe_checkout_session_id)
     VALUES (?, ?, ?, ?, 'pending', ?)`
  )
    .bind(orderId, userId, lotId, lot.website_price_cents, session.id)
    .run();

  return json({ checkoutUrl: session.url });
}

async function verifyStripeSignature(request, env, rawBody) {
  const sig = request.headers.get("stripe-signature");
  if (!sig) return false;

  const parts = Object.fromEntries(
    sig.split(",").map((p) => p.split("="))
  );
  const timestamp = parts.t;
  const signature = parts.v1;
  const signedPayload = `${timestamp}.${rawBody}`;

  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(env.STRIPE_WEBHOOK_SECRET),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sigBuffer = await crypto.subtle.sign(
    "HMAC",
    key,
    new TextEncoder().encode(signedPayload)
  );
  const expected = [...new Uint8Array(sigBuffer)]
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");

  return expected === signature;
}

async function handleStripeWebhook(request, env) {
  const rawBody = await request.text();
  const valid = await verifyStripeSignature(request, env, rawBody);
  if (!valid) return json({ error: "Invalid signature" }, 400);

  const event = JSON.parse(rawBody);

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;
    const orderId = session.metadata.order_id;
    const lotId = session.metadata.lot_id;

    await env.DB.batch([
      env.DB.prepare(
        "UPDATE orders SET status = 'paid', paid_at = datetime('now'), stripe_payment_intent_id = ? WHERE id = ?"
      ).bind(session.payment_intent, orderId),
      env.DB.prepare(
        "UPDATE lots SET status = 'sold', updated_at = datetime('now') WHERE id = ?"
      ).bind(lotId),
    ]);
  }

  return json({ received: true });
}

async function handleOrders(request, env) {
  const userId = await getAuthedUserId(request, env);
  if (!userId) return json({ error: "Please log in first" }, 401);

  const { results } = await env.DB.prepare(
    `SELECT o.id, o.amount_cents, o.status, o.created_at, o.paid_at, l.name AS lot_name
     FROM orders o JOIN lots l ON l.id = o.lot_id
     WHERE o.clerk_user_id = ?
     ORDER BY o.created_at DESC`
  )
    .bind(userId)
    .all();

  return json({ orders: results });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/lots" && request.method === "GET") {
      return handleListLots(env);
    }
    if (url.pathname === "/api/checkout" && request.method === "POST") {
      return handleCheckout(request, env);
    }
    if (url.pathname === "/api/webhook/stripe" && request.method === "POST") {
      return handleStripeWebhook(request, env);
    }
    if (url.pathname === "/api/orders" && request.method === "GET") {
      return handleOrders(request, env);
    }

    // Anything else: let the static assets binding handle it
    return env.ASSETS.fetch(request);
  },
};

// -----------------------------------------------------------------------
// One-time setup, in order:
//
// 1. Create the D1 database:
//      npx wrangler d1 create holly-hill-shop
//    Copy the database_id it prints into wrangler.jsonc (see updated file).
//
// 2. Load the schema:
//      npx wrangler d1 execute holly-hill-shop --file=schema.sql --remote
//
// 3. Install dependencies (this repo now needs a real build step):
//      npm install @clerk/backend
//
// 4. Set your secrets (never commit these to GitHub):
//      npx wrangler secret put STRIPE_SECRET_KEY
//      npx wrangler secret put STRIPE_WEBHOOK_SECRET
//      npx wrangler secret put CLERK_SECRET_KEY
//
// 5. In Cloudflare Pages build settings, set the Build command to:
//      npm install
//    (it was blank before -- now there's a real dependency to install)
//
// 6. In Stripe's dashboard, add a webhook endpoint pointing to:
//      https://hollyhillohio.com/api/webhook/stripe
//    listening for the "checkout.session.completed" event.
// -----------------------------------------------------------------------
