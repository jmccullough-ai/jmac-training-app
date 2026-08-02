JMAC TRAINING — STANDALONE PWA WRAPPER v1

PURPOSE 
This wrapper gives the phone app a real web manifest, Apple touch icon,
maskable Android icon, service worker, standalone display, install button,
share button, and custom JMAC logo.

CONNECTED APPS SCRIPT APP
https://script.google.com/a/macros/bwschools.net/s/AKfycbzIo2PjrNQPZ_tX1VvXRaD9XyvfluWTHlDI-dWa15VYQ7wg_UVq_1o87ifhcx3aE85IJA/exec

EASIEST DEPLOYMENT: GITHUB PAGES
1. Unzip this package.
2. Create a new PUBLIC GitHub repository, such as:
   jmac-training-app
3. Upload every file and folder from this package to the repository root.
4. Open the repository Settings.
5. Select Pages.
6. Under Build and deployment, choose:
   Source: Deploy from a branch
   Branch: main
   Folder: / (root)
7. Save.
8. GitHub will provide an HTTPS address similar to:
   https://YOUR-USERNAME.github.io/jmac-training-app/
9. Open that GitHub Pages address on the phone.
10. iPhone: Safari > Share > Add to Home Screen > Open as Web App.
11. Android: Chrome > Install app or use the Install button.

NETLIFY ALTERNATIVE
Drag the unzipped folder into Netlify Drop. Netlify will create an HTTPS link.

IMPORTANT
- Install the WRAPPER URL, not the script.google.com URL.
- The wrapper uses your existing Apps Script web app for the actual workout.
- The Google Sheet remains the data store.
- The Apps Script project must keep XFrameOptionsMode.ALLOWALL.
- Because the Apps Script link is tied to bwschools.net, the user may need
  to sign into that Google Workspace account.
- If Google sign-in is blocked inside the wrapper, tap the top-right
  Open Directly button, sign in, then return to the wrapper.

FILES
- index.html / 404.html
- manifest.webmanifest
- service-worker.js
- app-config.js
- offline.html
- icons/
- netlify.toml
- .nojekyll

TO CHANGE THE APPS SCRIPT URL LATER
Edit app-config.js and replace appUrl.
