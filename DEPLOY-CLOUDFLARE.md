# Deploy guide - abdurrahmanit.com

This folder is ready for a static Cloudflare Pages deployment.

## 1) GitHub repository
1. Make the repository Private if you do not want the source files public.
2. Upload everything in this folder to the repository root.
3. Keep `index.html` at the repository root.

## 2) Cloudflare Pages
1. Cloudflare Dashboard -> Workers & Pages -> Create application -> Pages -> Connect to Git.
2. Connect the GitHub repository.
3. Framework preset: None.
4. Build command: leave blank.
5. Build output directory: `/` (repository root). If the dashboard requires a value, use `.`.
6. Deploy.

After the Git integration is connected, later commits/pushes to the production branch will deploy automatically.

## 3) Custom domain
In the Pages project, add the custom domain:

`abdurrahmanit.com`

Cloudflare will create/guide the DNS record. The included `CNAME` file contains the same domain for compatibility, but Cloudflare Pages custom-domain configuration is controlled in the Cloudflare dashboard.

## 4) Cloudflare Access
Cloudflare Zero Trust -> Access -> Applications -> Add application -> Self-hosted.

Protect:
`abdurrahmanit.com/*`

Create an Allow policy for only the email addresses you approve. Use Email OTP or another identity provider you trust.

Do not create a broad Bypass policy for `Everyone`.

If you also use `www.abdurrahmanit.com`, add/protect that hostname too.

## 5) Important: block alternate public URLs
If the Pages project's `*.pages.dev` address remains publicly reachable, a visitor may bypass the custom-domain Access policy by opening that address directly. Protect the Pages project/preview URLs with Cloudflare Access as well, or disable access to alternate hostnames according to the options available in your Cloudflare dashboard.

Test while logged out/incognito:
- `https://abdurrahmanit.com`
- the project's `https://<project>.pages.dev` URL

Both should require authentication if both are intended to be private.

## 6) Secrets
Do not put passwords, GitHub Personal Access Tokens, Cloudflare API Tokens, or other secrets in this repository. This static site does not need those secrets for normal Git-connected deployment.

## Files
- `index.html` - portfolio home page
- `style.css` - portfolio styles
- `profile.png` - profile image
- `MD-Abdur-Rahman-CV.pdf` - downloadable CV
- `question-generator.html` - question generator UI
- `question-generator.css` - generator styles
- `question-generator.js` - generator logic
- `question-template.docx` - DOCX template
- `jszip.min.js` - local JSZip dependency
- `_headers` - basic security headers for Cloudflare Pages
