# TKR email open tracker

A 1x1 image served from your own domain. When a tracked email is opened, the image loads, this function
logs it to the master feed through the Apps Script web app (see apps-script/TKR-email-tracking.gs).

Setup (once):
1. Netlify > Add new site > Import from GitHub > this repo, branch apps-script-automations (or main once merged).
   Base directory: tracker. Publish directory: public. Functions: netlify/functions (netlify.toml sets these).
2. Site settings > Environment variables: TKR_SCRIPT_URL (the Apps Script web app /exec URL), TKR_SECRET (same value as the script property).
3. Domain management > add custom domain t.theknowledgeroom.uk. At your DNS host add the CNAME record Netlify shows
   (t -> <your-site>.netlify.app). Wait for the certificate.
4. Test: open https://t.theknowledgeroom.uk/o/test123456.gif in a browser. You should see a blank page and a 42-byte GIF.
