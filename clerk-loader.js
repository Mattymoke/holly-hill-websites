// Holly Hill Surplus — shared Clerk auth loader
//
// Edit ONLY these two lines once you have your real values from the
// Clerk dashboard (API keys page). Every shop page loads this same file,
// so you only need to update it in one place.

window.CLERK_PUBLISHABLE_KEY = "pk_test_Y2hhcm1lZC1iYXQtNzUuY2xlcmsuYWNjb3VudHMuZGV2JA"; // starts with pk_test_...
window.CLERK_FRONTEND_API = "charmed-bat-75.clerk.accounts.dev";

// -----------------------------------------------------------------------
// Everything below this line works automatically -- no need to touch it.
// -----------------------------------------------------------------------

window.clerkReady = new Promise(function (resolve, reject) {
  var script = document.createElement("script");
  script.setAttribute("data-clerk-publishable-key", window.CLERK_PUBLISHABLE_KEY);
  script.async = true;
  script.src = "https://" + window.CLERK_FRONTEND_API + "/npm/@clerk/clerk-js@latest/dist/clerk.browser.js";
  script.addEventListener("load", async function () {
    try {
      await window.Clerk.load();
      resolve(window.Clerk);
    } catch (err) {
      reject(err);
    }
  });
  script.addEventListener("error", function () {
    reject(new Error("Could not load Clerk. Check CLERK_FRONTEND_API in clerk-loader.js."));
  });
  document.head.appendChild(script);
});
