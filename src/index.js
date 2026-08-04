// Holly Hill Surplus — shop backend (Cloudflare Worker)
//
// Public routes:
//   GET  /api/lots              -> list available lots (public)
//   POST /api/checkout          -> create a Stripe Checkout session (requires login)
//   POST /api/webhook/stripe    -> Stripe calls this when payment completes
//   GET  /api/orders            -> order history for the logged-in buyer
//   GET  /media/:key            -> serves an uploaded lot photo from R2
//
// Admin-only routes (require the signed-in user to match ADMIN_USER_ID):
//   GET  /api/admin/lots         -> list ALL lots, any status
//   POST /api/admin/lots/create  -> create a new lot
//   POST /api/admin/lots/update  -> update an existing lot (price, status, etc.)
//   POST /api/admin/lots/delete  -> delete a lot
//   GET  /api/admin/orders       -> list all orders, across all buyers
//
// Script-only route (requires header "Authorization: Bearer <SYNC_API_KEY>",
// not Clerk auth -- this is for the Excel-to-website sync helper script):
//   POST /api/admin/sync-lot     -> upsert a lot + upload its photos to R2
//
// Requires these secrets:
//   wrangler secret put STRIPE_SECRET_KEY
//   wrangler secret put STRIPE_WEBHOOK_SECRET
//   wrangler secret put CLERK_SECRET_KEY
//   wrangler secret put ADMIN_USER_ID
//   wrangler secret put SYNC_API_KEY

import { verifyToken } from "@clerk/backend";

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json" },
  });
}

async function getAuthedUserId(request, env) {
  const authHeader = request.headers.get("Authorization") || "";
  const token = authHeader.replace(/^Bearer\s+/i, "");
  if (!token) return null;
  try {
    const payload = await verifyToken(token, {
      secretKey: env.CLERK_SECRET_KEY,
    });
    return payload.sub;
  } catch (err) {
    console.error("Clerk token verification failed:", err.message);
    return null;
  }
}

async function getAdminUserId(request, env) {
  const userId = await getAuthedUserId(request, env);
  if (!userId) return null;
  if (!env.ADMIN_USER_ID) {
    console.error("ADMIN_USER_ID is not configured -- refusing all admin requests");
    return null;
  }
  if (userId !== env.ADMIN_USER_ID) return null;
  return userId;
}

function isSyncAuthed(request, env) {
  const authHeader = request.headers.get("Authorization") || "";
  const token = authHeader.replace(/^Bearer\s+/i, "");
  if (!env.SYNC_API_KEY || !token) return false;
  return token === env.SYNC_API_KEY;
}

function sanitizeFilename(name) {
  return (name || "file").replace(/[^a-zA-Z0-9._-]/g, "_");
}

async function handleMedia(env, key) {
  const object = await env.IMAGES.get(key);
  if (!object) return new Response("Not found", { status: 404 });

  const headers = new Headers();
  headers.set("content-type", object.httpMetadata?.contentType || "application/octet-stream");
  headers.set("cache-control", "public, max-age=31536000, immutable");
  headers.set("etag", object.httpEtag);

  return new Response(object.body, { headers });
}

async function handleListLots(env) {
  const { results } = await env.DB.prepare(
    "SELECT id, name, category, description, website_price_cents, image_urls FROM lots WHERE status = 'available' ORDER BY created_at DESC"
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
  console.log("Webhook signature valid:", valid);
  if (!valid) return json({ error: "Invalid signature" }, 400);

  const event = JSON.parse(rawBody);
  console.log("Webhook event type:", event.type);

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;
    const orderId = session.metadata.order_id;
    const lotId = session.metadata.lot_id;
    console.log("Updating order:", orderId, "lot:", lotId);

    await env.DB.batch([
      env.DB.prepare(
        "UPDATE orders SET status = 'paid', paid_at = datetime('now'), stripe_payment_intent_id = ? WHERE id = ?"
      ).bind(session.payment_intent, orderId),
      env.DB.prepare(
        "UPDATE lots SET status = 'sold', updated_at = datetime('now') WHERE id = ?"
      ).bind(lotId),
    ]);
    console.log("Order update complete");
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

async function handleAdminListLots(request, env) {
  const adminId = await getAdminUserId(request, env);
  if (!adminId) return json({ error: "Not authorized" }, 403);

  const { results } = await env.DB.prepare(
    "SELECT * FROM lots ORDER BY created_at DESC"
  ).all();
  return json({ lots: results });
}

async function handleAdminCreateLot(request, env) {
  const adminId = await getAdminUserId(request, env);
  if (!adminId) return json({ error: "Not authorized" }, 403);

  const body = await request.json();
  const { id, name, category, description, website_price_cents } = body;

  if (!id || !name || !website_price_cents) {
    return json({ error: "id, name, and website_price_cents are required" }, 400);
  }

  try {
    await env.DB.prepare(
      `INSERT INTO lots (id, name, category, description, website_price_cents, status)
       VALUES (?, ?, ?, ?, ?, 'available')`
    )
      .bind(id, name, category || null, description || null, website_price_cents)
      .run();
  } catch (err) {
    return json({ error: "Could not create lot — that Lot # might already exist" }, 409);
  }

  return json({ success: true });
}

async function handleAdminUpdateLot(request, env) {
  const adminId = await getAdminUserId(request, env);
  if (!adminId) return json({ error: "Not authorized" }, 403);

  const body = await request.json();
  const { id, name, category, description, website_price_cents, status } = body;
  if (!id) return json({ error: "id is required" }, 400);

  await env.DB.prepare(
    `UPDATE lots SET
       name = COALESCE(?, name),
       category = COALESCE(?, category),
       description = COALESCE(?, description),
       website_price_cents = COALESCE(?, website_price_cents),
       status = COALESCE(?, status),
       updated_at = datetime('now')
     WHERE id = ?`
  )
    .bind(
      name ?? null,
      category ?? null,
      description ?? null,
      website_price_cents ?? null,
      status ?? null,
      id
    )
    .run();

  return json({ success: true });
}

async function handleAdminDeleteLot(request, env) {
  const adminId = await getAdminUserId(request, env);
  if (!adminId) return json({ error: "Not authorized" }, 403);

  const body = await request.json();
  const { id } = body;
  if (!id) return json({ error: "id is required" }, 400);

  await env.DB.prepare("DELETE FROM lots WHERE id = ?").bind(id).run();
  return json({ success: true });
}

async function handleAdminListOrders(request, env) {
  const adminId = await getAdminUserId(request, env);
  if (!adminId) return json({ error: "Not authorized" }, 403);

  const { results } = await env.DB.prepare(
    `SELECT o.id, o.clerk_user_id, o.amount_cents, o.status, o.created_at, o.paid_at, l.name AS lot_name
     FROM orders o JOIN lots l ON l.id = o.lot_id
     ORDER BY o.created_at DESC LIMIT 100`
  ).all();

  return json({ orders: results });
}

const SYNC_LOT_STATUSES = ["available", "reserved", "sold", "shipped"];

async function handleSyncLot(request, env) {
  if (!isSyncAuthed(request, env)) {
    return json({ error: "Not authorized" }, 401);
  }

  const formData = await request.formData();

  const id = (formData.get("id") || "").toString().trim();
  const name = (formData.get("name") || "").toString().trim();
  const category = formData.get("category") ? formData.get("category").toString().trim() : null;
  const description = formData.get("description") ? formData.get("description").toString().trim() : null;
  const status = (formData.get("status") || "").toString().trim();
  const websitePriceCents = Number(formData.get("website_price_cents"));
  const clearImages = ["1", "true", "yes"].includes(
    (formData.get("clear_images") || "").toString().trim().toLowerCase()
  );

  if (!id || !name) {
    return json({ error: "id and name are required" }, 400);
  }
  if (!Number.isInteger(websitePriceCents) || websitePriceCents < 0) {
    return json({ error: "website_price_cents must be a non-negative integer" }, 400);
  }
  if (!SYNC_LOT_STATUSES.includes(status)) {
    return json({ error: "status must be one of " + SYNC_LOT_STATUSES.join(", ") }, 400);
  }

  const files = formData.getAll("images").filter((f) => f instanceof File && f.size > 0);
  const newImageUrls = [];
  for (const file of files) {
    const key = `${id}-${crypto.randomUUID()}-${sanitizeFilename(file.name)}`;
    await env.IMAGES.put(key, file, {
      httpMetadata: { contentType: file.type || "application/octet-stream" },
    });
    newImageUrls.push(`https://hollyhillohio.com/media/${encodeURIComponent(key)}`);
  }

  const existing = await env.DB.prepare("SELECT image_urls FROM lots WHERE id = ?").bind(id).first();

  let finalImageUrls = newImageUrls;
  if (!clearImages && existing?.image_urls) {
    try {
      finalImageUrls = JSON.parse(existing.image_urls).concat(newImageUrls);
    } catch (err) {
      finalImageUrls = newImageUrls;
    }
  }

  if (existing) {
    await env.DB.prepare(
      `UPDATE lots SET
         name = ?, category = ?, description = ?, website_price_cents = ?,
         status = ?, image_urls = ?, updated_at = datetime('now')
       WHERE id = ?`
    )
      .bind(name, category, description, websitePriceCents, status, JSON.stringify(finalImageUrls), id)
      .run();
  } else {
    await env.DB.prepare(
      `INSERT INTO lots (id, name, category, description, website_price_cents, status, image_urls)
       VALUES (?, ?, ?, ?, ?, ?, ?)`
    )
      .bind(id, name, category, description, websitePriceCents, status, JSON.stringify(finalImageUrls))
      .run();
  }

  const updated = await env.DB.prepare("SELECT * FROM lots WHERE id = ?").bind(id).first();
  return json({
    lot: {
      ...updated,
      image_urls: updated.image_urls ? JSON.parse(updated.image_urls) : [],
    },
  });
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
    if (url.pathname.startsWith("/media/") && request.method === "GET") {
      const key = decodeURIComponent(url.pathname.slice("/media/".length));
      return handleMedia(env, key);
    }

    if (url.pathname === "/api/admin/lots" && request.method === "GET") {
      return handleAdminListLots(request, env);
    }
    if (url.pathname === "/api/admin/lots/create" && request.method === "POST") {
      return handleAdminCreateLot(request, env);
    }
    if (url.pathname === "/api/admin/lots/update" && request.method === "POST") {
      return handleAdminUpdateLot(request, env);
    }
    if (url.pathname === "/api/admin/lots/delete" && request.method === "POST") {
      return handleAdminDeleteLot(request, env);
    }
    if (url.pathname === "/api/admin/orders" && request.method === "GET") {
      return handleAdminListOrders(request, env);
    }
    if (url.pathname === "/api/admin/sync-lot" && request.method === "POST") {
      return handleSyncLot(request, env);
    }

    return env.ASSETS.fetch(request);
  },
};
