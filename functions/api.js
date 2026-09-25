/**
 * CrushOps API Proxy — Cloudflare Pages Functions
 * Route: POST /api → diteruskan ke Google Apps Script
 */

const GAS_URL = 'https://script.google.com/macros/s/AKfycbz94wl1Sg5SiEvnxd0UH5AFMNSSAla63BWHdHYPJ1bJOq0c9K2Pk7sTv4aWjIyqhSj0/exec';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Max-Age': '86400'
};

export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: CORS });
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.text();
    const gasRes = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: body,
      redirect: 'follow'
    });
    const data = await gasRes.text();
    return new Response(data, {
      status: 200,
      headers: {
        ...CORS,
        'Content-Type': 'application/json;charset=utf-8',
        'Cache-Control': 'no-store'
      }
    });
  } catch (err) {
    return new Response(
      JSON.stringify({ ok: false, error: String(err.message || err) }),
      { status: 500, headers: { ...CORS, 'Content-Type': 'application/json;charset=utf-8' } }
    );
  }
}

export async function onRequestGet() {
  return new Response(
    JSON.stringify({ ok: false, error: 'Gunakan POST untuk endpoint ini.' }),
    { status: 405, headers: { ...CORS, 'Content-Type': 'application/json' } }
  );
}
