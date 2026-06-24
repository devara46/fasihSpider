// background.js - Minimal service worker required for MV3
// The extension stores last results in chrome.storage.local so
// the collector page can resume after the service worker is killed.

chrome.runtime.onInstalled.addListener(() => {
  console.log('FASIH Region Collector installed.');
});
