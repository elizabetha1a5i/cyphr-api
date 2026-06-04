/**
 * Cyphr Proxy — Provider-Agnostic AI Worker
 * ==========================================
 * Supports: Anthropic (Claude), OpenAI (GPT), Google (Gemini)
 *
 * To switch provider: change MODEL_PROVIDER in your Worker environment variables.
 * Everything else stays the same — the frontend never needs to change.
 *
 * Environment variables to set in Cloudflare Worker → Settings → Variables:
 *
 *   MODEL_PROVIDER      = "anthropic"   ← change to "openai" or "gemini" to switch
 *   ANTHROPIC_API_KEY   = sk-ant-...    ← only needed if provider = anthropic
 *   OPENAI_API_KEY      = sk-...        ← only needed if provider = openai
 *   GEMINI_API_KEY      = AIza...       ← only needed if provider = gemini
 *
 *   MODEL_OVERRIDE      = (optional) override the default model for the active provider
 *                         e.g. "claude-haiku-4-5-20251001" or "gpt-4o-mini" or "gemini-2.0-flash"
 *
 * The worker also handles the Samsung report upload metadata via KV (unchanged).
 */

// ── Default models per provider ──────────────────────────────────────────────
const DEFAULT_MODELS = {
  anthropic: 'claude-haiku-4-5-20251001',
  openai:    'gpt-4o-mini',
  gemini:    'gemini-2.0-flash',
};

// ── CORS headers ──────────────────────────────────────────────────────────────
const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

function corsResponse(body, status = 200, extra = {}) {
  return new Response(body, {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS, ...extra },
  });
}


// ── Main handler ──────────────────────────────────────────────────────────────
export default {
  async fetch(request, env) {

    // Preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS });
    }

    const url = new URL(request.url);

    // ── Samsung report upload metadata (KV) — unchanged from original ─────────
    if (url.pathname === '/report-upload') {
      return handleReportUpload(request, env);
    }
    if (url.pathname === '/report-status') {
      return handleReportStatus(request, env);
    }

    // ── Health check ──────────────────────────────────────────────────────────
    if (url.pathname === '/health') {
      const provider = env.MODEL_PROVIDER || 'anthropic';
      const model    = env.MODEL_OVERRIDE || DEFAULT_MODELS[provider] || 'unknown';
      return corsResponse(JSON.stringify({ status: 'ok', provider, model }));
    }

    // ── AI proxy ──────────────────────────────────────────────────────────────
    if (request.method !== 'POST') {
      return corsResponse(JSON.stringify({ error: 'Method not allowed' }), 405);
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return corsResponse(JSON.stringify({ error: 'Invalid JSON body' }), 400);
    }

    const provider = (env.MODEL_PROVIDER || 'anthropic').toLowerCase();
    const model    = env.MODEL_OVERRIDE || DEFAULT_MODELS[provider];

    try {
      switch (provider) {
        case 'anthropic': return await callAnthropic(body, model, env);
        case 'openai':    return await callOpenAI(body, model, env);
        case 'gemini':    return await callGemini(body, model, env);
        default:
          return corsResponse(
            JSON.stringify({ error: `Unknown provider "${provider}". Use anthropic, openai, or gemini.` }),
            400
          );
      }
    } catch (err) {
      return corsResponse(
        JSON.stringify({ error: 'Provider call failed', detail: err.message }),
        502
      );
    }
  }
};


// ── Anthropic ─────────────────────────────────────────────────────────────────
async function callAnthropic(body, model, env) {
  if (!env.ANTHROPIC_API_KEY) {
    return corsResponse(JSON.stringify({ error: 'ANTHROPIC_API_KEY not set' }), 500);
  }

  // Normalise: accept either {messages, system, max_tokens} directly,
  // or the frontend's {prompt} shorthand
  const payload = normaliseToAnthropic(body, model);

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type':    'application/json',
      'x-api-key':       env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify(payload),
  });

  const data = await res.json();

  // Normalise response to common format before returning
  return corsResponse(JSON.stringify(normaliseFromAnthropic(data)));
}


// ── OpenAI ────────────────────────────────────────────────────────────────────
async function callOpenAI(body, model, env) {
  if (!env.OPENAI_API_KEY) {
    return corsResponse(JSON.stringify({ error: 'OPENAI_API_KEY not set' }), 500);
  }

  const payload = normaliseToOpenAI(body, model);

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': `Bearer ${env.OPENAI_API_KEY}`,
    },
    body: JSON.stringify(payload),
  });

  const data = await res.json();
  return corsResponse(JSON.stringify(normaliseFromOpenAI(data)));
}


// ── Gemini ────────────────────────────────────────────────────────────────────
async function callGemini(body, model, env) {
  if (!env.GEMINI_API_KEY) {
    return corsResponse(JSON.stringify({ error: 'GEMINI_API_KEY not set' }), 500);
  }

  const payload = normaliseToGemini(body);

  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${env.GEMINI_API_KEY}`,
    {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload),
    }
  );

  const data = await res.json();
  return corsResponse(JSON.stringify(normaliseFromGemini(data)));
}


// ── Normalisation: request → provider format ──────────────────────────────────
//
// The frontend always sends:
//   { prompt: "...", system: "...", max_tokens: 1000 }
//   OR the native Anthropic format directly (messages array)
//
// These functions translate to/from each provider's native format.
// The response is always normalised back to Anthropic's format
// so the frontend code never needs to change.

function normaliseToAnthropic(body, model) {
  // If already in Anthropic format, just ensure model is set
  if (body.messages) {
    return { max_tokens: 1000, ...body, model };
  }
  // Shorthand {prompt, system}
  const msg = { role: 'user', content: body.prompt || '' };
  const out = { model, max_tokens: body.max_tokens || 1000, messages: [msg] };
  if (body.system) out.system = body.system;
  return out;
}

function normaliseToOpenAI(body, model) {
  const messages = [];
  const system = body.system || (body.messages && body.messages.find(m => m.role === 'system')?.content);
  if (system) messages.push({ role: 'system', content: system });

  if (body.messages) {
    body.messages.filter(m => m.role !== 'system').forEach(m => messages.push(m));
  } else {
    messages.push({ role: 'user', content: body.prompt || '' });
  }

  return {
    model,
    max_tokens: body.max_tokens || 1000,
    messages,
  };
}

function normaliseToGemini(body) {
  const parts = [];
  if (body.system) parts.push({ text: `System: ${body.system}\n\n` });

  if (body.messages) {
    body.messages.forEach(m => {
      if (m.role !== 'system') parts.push({ text: m.content });
    });
  } else {
    parts.push({ text: body.prompt || '' });
  }

  return {
    contents: [{ parts }],
    generationConfig: { maxOutputTokens: body.max_tokens || 1000 },
  };
}


// ── Normalisation: provider response → Anthropic format ──────────────────────
//
// Always returns:
//   { content: [{ type: 'text', text: '...' }] }
//
// This matches what the frontend already expects from Anthropic.
// So d.content[0].text always works regardless of which provider ran.

function normaliseFromAnthropic(data) {
  // Already in the right format — pass through
  // Add a _provider field for debugging
  return { ...data, _provider: 'anthropic' };
}

function normaliseFromOpenAI(data) {
  if (data.error) {
    return { error: data.error, _provider: 'openai' };
  }
  const text = data.choices?.[0]?.message?.content || '';
  return {
    content: [{ type: 'text', text }],
    model:   data.model,
    usage:   data.usage,
    _provider: 'openai',
  };
}

function normaliseFromGemini(data) {
  if (data.error) {
    return { error: data.error, _provider: 'gemini' };
  }
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  return {
    content: [{ type: 'text', text }],
    _provider: 'gemini',
  };
}


// ── Samsung report upload (KV) — unchanged ────────────────────────────────────
async function handleReportUpload(request, env) {
  if (request.method !== 'POST') {
    return corsResponse(JSON.stringify({ error: 'POST only' }), 405);
  }
  try {
    const body = await request.json();
    const entry = {
      uploadedBy: body.uploadedBy || 'unknown',
      uploadedAt: new Date().toISOString(),
      fileName:   body.fileName   || 'unknown',
      weekLabel:  body.weekLabel  || 'unknown',
    };
    await env.CYPHR_SHARED.put('samsung_report_latest', JSON.stringify(entry));
    return corsResponse(JSON.stringify({ success: true, entry }));
  } catch (err) {
    return corsResponse(JSON.stringify({ error: err.message }), 500);
  }
}

async function handleReportStatus(request, env) {
  try {
    const raw = await env.CYPHR_SHARED.get('samsung_report_latest');
    if (!raw) return corsResponse(JSON.stringify({ uploaded: false }));
    return corsResponse(JSON.stringify({ uploaded: true, ...JSON.parse(raw) }));
  } catch (err) {
    return corsResponse(JSON.stringify({ error: err.message }), 500);
  }
}
