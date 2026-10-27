/* Little Bloom service worker: makes the app installable and receives FCM pushes while it is closed.
   The Firebase config is passed in the registration URL by index.html, so it lives in one place. */
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');

var config = {};
try { config = JSON.parse(new URL(self.location.href).searchParams.get('config') || '{}'); } catch (e) {}

if (config.apiKey) {
  firebase.initializeApp(config);
  // Pushes sent with a `notification` block are shown automatically by the SDK.
  firebase.messaging();
}

self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (event) { event.waitUntil(self.clients.claim()); });

// Local notifications (focus timer) open or focus the app when tapped.
self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  var target = new URL('./', self.location.href).href;
  event.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
    for (var i = 0; i < list.length; i++) {
      if (list[i].url.indexOf(target) === 0 && 'focus' in list[i]) return list[i].focus();
    }
    return self.clients.openWindow(target);
  }));
});
