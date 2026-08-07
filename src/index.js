// Holly Hill Surplus — shop backend (Cloudflare Worker)
//
// Public routes:
//   GET  /api/lots              -> list available lots (public)
//   GET  /api/lots/sold         -> recently sold lots, for the "Recently Sold" trust section (public)
//   GET  /api/lots/sold/count   -> total count of sold lots, for the shop stats bar (public)
//   GET  /api/lots/featured     -> top-profit featured lots, for the public "Featured Lots" section (public)
//   GET  /api/lots/:id          -> single lot detail, any status (public)
//   POST /api/checkout          -> create a Stripe Checkout session (requires login)
//   POST /api/webhook/stripe    -> Stripe calls this when payment completes
//   GET  /api/orders            -> order history for the logged-in buyer
//   POST /api/contact           -> contact form submission, sent via Resend
//   GET  /media/:key            -> serves an uploaded lot photo from R2
//
// Admin-only routes (require the signed-in user to match ADMIN_USER_ID):
//   GET  /api/admin/lots         -> list ALL lots, any status
//   POST /api/admin/lots/create  -> create a new lot
//   POST /api/admin/lots/update  -> update an existing lot (price, status, etc.)
//   POST /api/admin/lots/delete  -> delete a lot
//   GET  /api/admin/orders       -> list all orders, across all buyers
//
// Script-only routes (require header "Authorization: Bearer <SYNC_API_KEY>",
// not Clerk auth -- these are for the Excel-to-website sync helper scripts):
//   POST /api/admin/sync-lot     -> upsert a lot + upload its photos to R2
//   GET  /api/sync/lots          -> id/status/updated_at for every lot, for two-way sync
//   GET  /api/sync/orders        -> orders + buyer/lot info, for the read-only Sold Lots tracker
//
// Requires these secrets:
//   wrangler secret put STRIPE_SECRET_KEY
//   wrangler secret put STRIPE_WEBHOOK_SECRET
//   wrangler secret put CLERK_SECRET_KEY
//   wrangler secret put ADMIN_USER_ID
//   wrangler secret put SYNC_API_KEY
//   wrangler secret put RESEND_API_KEY
//   wrangler secret put TURNSTILE_SECRET_KEY

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

async function uploadLotImages(env, id, files) {
  const newImageUrls = [];
  for (const file of files) {
    const key = `${id}-${crypto.randomUUID()}-${sanitizeFilename(file.name)}`;
    await env.IMAGES.put(key, file, {
      httpMetadata: { contentType: file.type || "application/octet-stream" },
    });
    newImageUrls.push(`https://hollyhillohio.com/media/${encodeURIComponent(key)}`);
  }
  return newImageUrls;
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
    "SELECT id, name, category, description, website_price_cents, image_urls, created_at FROM lots WHERE status = 'available' ORDER BY created_at DESC"
  ).all();
  return json({ lots: results });
}

async function handleSoldCount(env) {
  const row = await env.DB.prepare(
    "SELECT COUNT(*) AS count FROM lots WHERE status = 'sold'"
  ).first();
  return json({ count: row ? row.count : 0 });
}

async function handleListSoldLots(env) {
  const { results } = await env.DB.prepare(
    "SELECT id, name, category, website_price_cents, image_urls, updated_at FROM lots WHERE status = 'sold' ORDER BY updated_at DESC LIMIT 8"
  ).all();
  return json({ lots: results });
}

async function handleListFeaturedLots(env) {
  // Never select or expose original cost / profit here -- those figures
  // don't exist in D1 at all, only in Excel (Original Price / Profit are
  // PRIVATE columns that never sync).
  const { results } = await env.DB.prepare(
    "SELECT id, name, category, website_price_cents, image_urls FROM lots WHERE status = 'available' AND is_featured = 1 ORDER BY updated_at DESC LIMIT 6"
  ).all();
  return json({ lots: results });
}

async function handleLotDetail(env, id) {
  const lot = await env.DB.prepare(
    "SELECT id, name, category, description, website_price_cents, status, image_urls FROM lots WHERE id = ?"
  )
    .bind(id)
    .first();

  if (!lot) return json({ error: "Lot not found" }, 404);

  return json({ lot });
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

// Fire-and-forget from handleStripeWebhook via ctx.waitUntil -- must never
// throw, since a Resend failure here must never affect the webhook response
// or the order/lot status update that already happened before this runs.
async function sendOrderConfirmationEmail(env, session, orderId, lotId) {
  try {
    const buyerEmail = session.customer_details && session.customer_details.email;
    if (!buyerEmail) {
      console.error("Order confirmation email skipped -- no buyer email on session for order", orderId);
      return;
    }

    const lot = await env.DB.prepare("SELECT name, website_price_cents FROM lots WHERE id = ?")
      .bind(lotId)
      .first();
    const lotName = lot ? lot.name : "your item";
    const amount = lot ? "$" + (lot.website_price_cents / 100).toFixed(2) : "";

    const html =
      `<p>Thanks for your order from Holly Hill Surplus!</p>` +
      `<p><strong>Item:</strong> ${escapeHtml(lotName)}</p>` +
      (amount ? `<p><strong>Amount paid:</strong> ${escapeHtml(amount)}</p>` : "") +
      `<p><strong>Order ID:</strong> ${escapeHtml(orderId)}</p>` +
      `<p>We'll be in touch about pickup or shipping. Just reply to this email with any questions.</p>`;

    const resendRes = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + env.RESEND_API_KEY,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: "Holly Hill Surplus <orders@hollyhillohio.com>",
        to: buyerEmail,
        subject: "Order confirmed — " + lotName,
        html,
      }),
    });

    if (!resendRes.ok) {
      const errText = await resendRes.text();
      console.error("Order confirmation email failed:", resendRes.status, errText);
    }
  } catch (err) {
    console.error("Order confirmation email threw:", err.message);
  }
}

async function handleStripeWebhook(request, env, ctx) {
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
    const buyerEmail = (session.customer_details && session.customer_details.email) || null;
    console.log("Updating order:", orderId, "lot:", lotId);

    await env.DB.batch([
      env.DB.prepare(
        "UPDATE orders SET status = 'paid', paid_at = datetime('now'), stripe_payment_intent_id = ?, buyer_email = ? WHERE id = ?"
      ).bind(session.payment_intent, buyerEmail, orderId),
      env.DB.prepare(
        "UPDATE lots SET status = 'sold', updated_at = datetime('now') WHERE id = ?"
      ).bind(lotId),
    ]);
    console.log("Order update complete");

    // Non-blocking: the order is already marked paid above, so a Resend
    // failure here can never affect order processing or the response
    // below. waitUntil (not a bare unawaited call) so the runtime doesn't
    // kill the request in-flight once the response is returned.
    if (ctx && ctx.waitUntil) {
      ctx.waitUntil(sendOrderConfirmationEmail(env, session, orderId, lotId));
    }
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

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function verifyTurnstile(token, env) {
  const params = new URLSearchParams({
    secret: env.TURNSTILE_SECRET_KEY,
    response: token,
  });
  const res = await fetch("https://challenges.cloudflare.com/turnstile/v0/siteverify", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
  });
  const data = await res.json();
  return data.success === true;
}

async function handleContact(request, env) {
  let body;
  try {
    body = await request.json();
  } catch (err) {
    return json({ error: "Invalid request body" }, 400);
  }

  const { name, email, message, turnstileToken, honeypot } = body || {};

  // A bot filled in a field real users never see -- silently pretend to
  // succeed rather than tip it off that it was caught.
  if (honeypot) {
    return json({ success: true });
  }

  if (!turnstileToken) {
    return json({ error: "Please complete the verification check." }, 400);
  }
  const turnstileOk = await verifyTurnstile(turnstileToken, env);
  if (!turnstileOk) {
    return json({ error: "Verification check failed. Please try again." }, 400);
  }

  const trimmedName = (name || "").toString().trim();
  const trimmedEmail = (email || "").toString().trim();
  const trimmedMessage = (message || "").toString().trim();

  const fieldErrors = {};
  if (!trimmedName) fieldErrors.name = "Name is required.";
  if (!EMAIL_REGEX.test(trimmedEmail)) fieldErrors.email = "Please enter a valid email address.";
  if (!trimmedMessage) {
    fieldErrors.message = "Message is required.";
  } else if (trimmedMessage.length > 3000) {
    fieldErrors.message = "Message must be under 3000 characters.";
  }

  if (Object.keys(fieldErrors).length > 0) {
    return json({ error: "Please fix the errors below.", fields: fieldErrors }, 400);
  }

  const html =
    `<p><strong>Name:</strong> ${escapeHtml(trimmedName)}</p>` +
    `<p><strong>Email:</strong> ${escapeHtml(trimmedEmail)}</p>` +
    `<p><strong>Message:</strong></p>` +
    `<p>${escapeHtml(trimmedMessage).replace(/\n/g, "<br>")}</p>`;

  const resendRes = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + env.RESEND_API_KEY,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      from: "Holly Hill Contact Form <contact@hollyhillohio.com>",
      to: "surplus@hollyhillohio.com",
      reply_to: trimmedEmail,
      subject: "New contact form message from " + trimmedName,
      html,
    }),
  });

  if (!resendRes.ok) {
    const errText = await resendRes.text();
    console.error("Resend send failed:", resendRes.status, errText);
    return json({ error: "Could not send your message right now. Please try again later." }, 502);
  }

  return json({ success: true });
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

  const formData = await request.formData();
  const id = (formData.get("id") || "").toString().trim();
  const name = (formData.get("name") || "").toString().trim();
  const category = formData.get("category") ? formData.get("category").toString().trim() : null;
  const description = formData.get("description") ? formData.get("description").toString().trim() : null;
  const websitePriceCents = Number(formData.get("website_price_cents"));

  if (!id || !name || !Number.isInteger(websitePriceCents) || websitePriceCents < 0) {
    return json({ error: "id, name, and website_price_cents are required" }, 400);
  }

  const files = formData.getAll("images").filter((f) => f instanceof File && f.size > 0);
  const imageUrls = await uploadLotImages(env, id, files);

  try {
    await env.DB.prepare(
      `INSERT INTO lots (id, name, category, description, website_price_cents, status, image_urls)
       VALUES (?, ?, ?, ?, ?, 'available', ?)`
    )
      .bind(id, name, category, description, websitePriceCents, JSON.stringify(imageUrls))
      .run();
  } catch (err) {
    return json({ error: "Could not create lot — that Lot # might already exist" }, 409);
  }

  return json({ success: true });
}

async function handleAdminUpdateLot(request, env) {
  const adminId = await getAdminUserId(request, env);
  if (!adminId) return json({ error: "Not authorized" }, 403);

  const formData = await request.formData();
  const id = (formData.get("id") || "").toString().trim();
  if (!id) return json({ error: "id is required" }, 400);

  const name = formData.get("name") ? formData.get("name").toString().trim() : null;
  const category = formData.get("category") ? formData.get("category").toString().trim() : null;
  const description = formData.get("description") ? formData.get("description").toString().trim() : null;
  const websitePriceCentsRaw = formData.get("website_price_cents");
  const websitePriceCents =
    websitePriceCentsRaw !== null && websitePriceCentsRaw !== "" ? Number(websitePriceCentsRaw) : null;
  const status = formData.get("status") ? formData.get("status").toString().trim() : null;

  // Photos append to whatever's already on the lot -- same behavior as
  // the Excel sync-lot endpoint. Only touched when files are actually
  // provided, so a plain price/status save doesn't disturb image_urls.
  const files = formData.getAll("images").filter((f) => f instanceof File && f.size > 0);
  let imageUrlsUpdate = null;
  if (files.length) {
    const newImageUrls = await uploadLotImages(env, id, files);
    const existing = await env.DB.prepare("SELECT image_urls FROM lots WHERE id = ?").bind(id).first();
    let finalImageUrls = newImageUrls;
    if (existing?.image_urls) {
      try {
        finalImageUrls = JSON.parse(existing.image_urls).concat(newImageUrls);
      } catch (err) {
        finalImageUrls = newImageUrls;
      }
    }
    imageUrlsUpdate = JSON.stringify(finalImageUrls);
  }

  await env.DB.prepare(
    `UPDATE lots SET
       name = COALESCE(?, name),
       category = COALESCE(?, category),
       description = COALESCE(?, description),
       website_price_cents = COALESCE(?, website_price_cents),
       status = COALESCE(?, status),
       image_urls = COALESCE(?, image_urls),
       updated_at = datetime('now')
     WHERE id = ?`
  )
    .bind(
      name,
      category,
      description,
      websitePriceCents,
      status,
      imageUrlsUpdate,
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
  const isFeatured = ["1", "true", "yes"].includes(
    (formData.get("featured") || "").toString().trim().toLowerCase()
  )
    ? 1
    : 0;

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
  const newImageUrls = await uploadLotImages(env, id, files);

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
         status = ?, image_urls = ?, is_featured = ?, updated_at = datetime('now')
       WHERE id = ?`
    )
      .bind(name, category, description, websitePriceCents, status, JSON.stringify(finalImageUrls), isFeatured, id)
      .run();
  } else {
    await env.DB.prepare(
      `INSERT INTO lots (id, name, category, description, website_price_cents, status, image_urls, is_featured)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    )
      .bind(id, name, category, description, websitePriceCents, status, JSON.stringify(finalImageUrls), isFeatured)
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

async function handleSyncLotsList(request, env) {
  if (!isSyncAuthed(request, env)) {
    return json({ error: "Not authorized" }, 401);
  }

  const { results } = await env.DB.prepare(
    "SELECT id, status, updated_at FROM lots"
  ).all();
  return json({ lots: results });
}

async function handleSyncOrders(request, env) {
  if (!isSyncAuthed(request, env)) {
    return json({ error: "Not authorized" }, 401);
  }

  const { results } = await env.DB.prepare(
    `SELECT o.id, o.lot_id, l.name AS lot_name, l.category, o.amount_cents,
            o.buyer_email, o.status, o.created_at, o.paid_at
     FROM orders o JOIN lots l ON l.id = o.lot_id
     ORDER BY o.created_at DESC`
  ).all();

  return json({ orders: results });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname === "/api/lots" && request.method === "GET") {
      return handleListLots(env);
    }
    if (url.pathname === "/api/lots/sold/count" && request.method === "GET") {
      return handleSoldCount(env);
    }
    if (url.pathname === "/api/lots/sold" && request.method === "GET") {
      return handleListSoldLots(env);
    }
    if (url.pathname === "/api/lots/featured" && request.method === "GET") {
      return handleListFeaturedLots(env);
    }
    if (url.pathname.startsWith("/api/lots/") && request.method === "GET") {
      const id = decodeURIComponent(url.pathname.slice("/api/lots/".length));
      return handleLotDetail(env, id);
    }
    if (url.pathname === "/api/checkout" && request.method === "POST") {
      return handleCheckout(request, env);
    }
    if (url.pathname === "/api/webhook/stripe" && request.method === "POST") {
      return handleStripeWebhook(request, env, ctx);
    }
    if (url.pathname === "/api/orders" && request.method === "GET") {
      return handleOrders(request, env);
    }
    if (url.pathname === "/api/contact" && request.method === "POST") {
      return handleContact(request, env);
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
    if (url.pathname === "/api/sync/lots" && request.method === "GET") {
      return handleSyncLotsList(request, env);
    }
    if (url.pathname === "/api/sync/orders" && request.method === "GET") {
      return handleSyncOrders(request, env);
    }

    return env.ASSETS.fetch(request);
  },
};
