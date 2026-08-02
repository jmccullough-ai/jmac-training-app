JMAC TRAINING PWA — DIRECT LAUNCH FIX v2

WHY THE SCREEN WAS BLANK
The Apps Script app was being loaded inside an iframe. On iPhone, Google Workspace
authentication and cross-site cookie restrictions can block that embedded page,
leaving a blank white area.

THIS BUILD FIXES IT BY
- Removing the iframe.
- Opening the Apps Script training app directly in the same window.
- Preserving the installed JMAC Home Screen icon from the GitHub Pages wrapper.
- Automatically opening the workout when launched from the installed app.
- Keeping manual Open Training App, Share, Install, and Copy Link controls.

UPDATE GITHUB
1. Unzip this package.
2. In the GitHub repository, upload and replace ALL existing files.
3. Commit with:
   Direct launch fix v2
4. Wait for the Pages deployment to finish.
5. Delete the old JMAC Home Screen app.
6. Open the GitHub Pages URL in Safari.
7. Refresh once.
8. Share > Add to Home Screen > Open as Web App > Add.
9. Open the new JMAC icon.

The installed icon belongs to the GitHub Pages wrapper. The wrapper then launches
the existing Apps Script app directly, avoiding the blank embedded screen.
