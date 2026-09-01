// Bray Fitness Tracker — CLEAN service worker
// - Caches app shell
// - Always fetches Google Script live
// - Ignores browser extension traffic
// - Prevents console spam from failed extension requests

const CACHE_NAME = "bray-fitness-v1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // 1. Ignore browser extension traffic (DeepSeek, StudyQuicks, Autotrack, etc.)
  if (
    url.protocol === "chrome-extension:" ||
    url.hostname.includes("studyquicks") ||
    url.hostname.includes("autotrack") ||
    url.hostname.includes("math_h5")
  ) {
    return; // Do nothing — prevents spam
  }

  // 2. Always fetch Google Apps Script live (never cache)
  if (url.hostname.includes("script.google.com")) {
    event.respondWith(
      fetch(event.request).
