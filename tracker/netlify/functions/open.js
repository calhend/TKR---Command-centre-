// Serves a 1x1 transparent GIF and logs the open to the TKR master feed via the Apps Script web app.
// Environment variables (Netlify site settings): TKR_SCRIPT_URL (the Apps Script web app /exec URL), TKR_SECRET.
const GIF = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');

exports.handler = async (event) => {
  const m = (event.path || '').match(/\/o\/([A-Za-z0-9_-]{6,40})(?:\.gif)?$/);
  const id = m ? m[1] : '';
  if (id && process.env.TKR_SCRIPT_URL && process.env.TKR_SECRET) {
    const ua = (event.headers && (event.headers['user-agent'] || event.headers['User-Agent'])) || '';
    const url = process.env.TKR_SCRIPT_URL + '?k=' + encodeURIComponent(process.env.TKR_SECRET) +
      '&id=' + encodeURIComponent(id) + '&e=open&ua=' + encodeURIComponent(ua.slice(0, 200));
    try { await Promise.race([fetch(url, { redirect: 'follow' }), new Promise(r => setTimeout(r, 4000))]); } catch (e) { /* never break the image */ }
  }
  return {
    statusCode: 200,
    isBase64Encoded: true,
    headers: { 'Content-Type': 'image/gif', 'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0', 'Content-Length': String(GIF.length) },
    body: GIF.toString('base64'),
  };
};
