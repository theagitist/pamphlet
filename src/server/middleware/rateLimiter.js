import rateLimit, { ipKeyGenerator } from 'express-rate-limit';

// Use Cloudflare's real client IP when behind their proxy,
// fall back to the standard IP key generator for direct connections / local dev.
function keyGenerator(req, res) {
  return req.headers['cf-connecting-ip'] || ipKeyGenerator(req, res);
}

// Upload: 2 requests per IP per minute
export const uploadLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 2,
  keyGenerator,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many uploads. Please wait a minute before trying again.' },
});

// Download: 1 request per 5 seconds per IP
export const downloadLimiter = rateLimit({
  windowMs: 5 * 1000,
  max: 1,
  keyGenerator,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Please wait a few seconds before downloading again.' },
});

// Status polling: 60 requests per minute per IP (generous but prevents abuse)
export const statusLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 60,
  keyGenerator,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many status requests.' },
});
