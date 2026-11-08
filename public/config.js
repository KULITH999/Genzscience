// Paste your Google Apps Script Web App URL here (Deploy > New deployment > Web app)
// Example: https://script.google.com/macros/s/AKfycb.../exec
window.API_URL = 'https://script.google.com/macros/s/AKfycbyEoSWLOV0czspM-862X6r40IK3E1igOzq5ymP2KlfLa8JZIVRHKHlgtwUYr29O2qly/exec';

// Universal API helper
window.callApi = async function (action, ...args) {
  const url = window.API_URL;
  if (!url || url.indexOf('PASTE_YOUR') === 0) {
    throw new Error('API URL not configured. Edit public/config.js first.');
  }
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: action, args: args })
  });
  if (!res.ok) throw new Error('Network error: ' + res.status);
  return res.json();
};