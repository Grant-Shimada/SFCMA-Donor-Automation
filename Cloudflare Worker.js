// 1. Define list of permitted domains
const allowedOrigins = [
  'https://sfcma.squarespace.com',
  'https://sfcivicmusic.org',
  'https://www.sfcivicmusic.org' // Added www variant just in case
];

const RATE_LIMIT = 60;
const RATE_WINDOW_SECONDS = 60;

export default {
  async fetch(request, env) {
    // 2. Extract the incoming origin from the browser request
    const requestOrigin = request.headers.get('Origin');
    
    // 3. Check if the origin is allowed; if not, default to the main domain
    const originToReturn = allowedOrigins.includes(requestOrigin) 
      ? requestOrigin 
      : allowedOrigins[0];

    // 4. Generate dynamic CORS headers based on the current origin
    const corsHeaders = {
      'Access-Control-Allow-Origin': originToReturn,
      'Access-Control-Allow-Methods': 'HEAD, GET, PUT, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, X-API-Key',
    };

    const url = new URL(request.url);
    const key = url.pathname.slice(1);

    if (!key) return new Response("No file specified", { status: 400, headers: corsHeaders });

    // Handle Preflight Request
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    // ── Rate limiting ──────────────────────────────────────────
    // Skip rate limiting if no KV namespace is bound (fails safe).
    if (env.RATE_LIMIT) {
      try {
        const ip = request.headers.get('CF-Connecting-IP') || 'unknown';
        const rateKey = `rl:${ip}`;
        const current = parseInt((await env.RATE_LIMIT.get(rateKey)) || '0', 10);

        if (current >= RATE_LIMIT) {
          return new Response(
            JSON.stringify({ error: 'Too many requests. Please try again later.' }),
            {
              status: 429,
              headers: {
                ...corsHeaders,
                'Content-Type': 'application/json',
                'Retry-After': String(RATE_WINDOW_SECONDS),
              },
            }
          );
        }

        // Increment counter; set TTL on first request in the window
        if (current === 0) {
          await env.RATE_LIMIT.put(rateKey, '1', {
            expirationTtl: RATE_WINDOW_SECONDS,
          });
        } else {
          await env.RATE_LIMIT.put(rateKey, String(current + 1), {
            expirationTtl: RATE_WINDOW_SECONDS,
          });
        }
      } catch(err) {
        console.error('Rate limit KV error:', err);
      }
    }

    // Handle Data Storage (PUT)
    if (request.method === 'PUT') {
      if (request.headers.get('X-API-Key') !== env.API_KEY) {
        return new Response('Unauthorized', { status: 401, headers: corsHeaders });
      }
      
      const data = await request.text();
      await env["SFCMA-DONORS"].put(key, data);
      return new Response('Saved', { status: 200, headers: corsHeaders });
    }
    
    // Handle Metadata Retrieval (HEAD)
    if (request.method === 'HEAD') {
        const metadata = await env["SFCMA-DONORS"].head(key);
        if (!metadata) return new Response("File not found", { status: 404, headers: corsHeaders });
        return new Response(null, {
          headers: {
            ...corsHeaders,
            "Last-Modified": metadata.uploaded.toISOString()
          }
        });
    }

     // Handle Data Retrieval (GET)
    if (request.method === 'GET') {
      const data = await env["SFCMA-DONORS"].get(key);

      if (!data) {
        return new Response('File not found', { status: 404, headers: corsHeaders });
      }

      return new Response(data.body, { 
        headers: {
          ...corsHeaders,
          "Last-Modified": data.uploaded.toISOString()
        }
      });
    }

    // Catch-all
    return new Response('Method not allowed', { status: 405, headers: corsHeaders });
  }
};
