export const verifyTurnstile = async (req, res, next) => {
  const secretKey = process.env.TURNSTILE_SECRET_KEY;
  
  // Skip if not configured (development/testing)
  if (!secretKey) {
    return next();
  }

  const token = req.body['cf-turnstile-response'] || req.headers['x-turnstile-token'];

  if (!token) {
    return res.status(401).json({ error: 'Security verification required' });
  }

  try {
    const formData = new URLSearchParams();
    formData.append('secret', secretKey);
    formData.append('response', token);
    formData.append('remoteip', req.headers['cf-connecting-ip'] || req.ip);

    const result = await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      body: formData,
      method: 'POST',
    });

    const outcome = await result.json();

    if (outcome.success) {
      next();
    } else {
      res.status(403).json({ error: 'Security verification failed' });
    }
  } catch (err) {
    console.error('Turnstile verification error:', err);
    res.status(500).json({ error: 'Internal verification error' });
  }
};
