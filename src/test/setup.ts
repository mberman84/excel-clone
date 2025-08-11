import '@testing-library/jest-dom/vitest';
import { afterEach } from 'vitest';

// Set fixed window dimensions for consistent testing
Object.defineProperty(window, 'innerWidth', {
  writable: true,
  configurable: true,
  value: 1200,
});

Object.defineProperty(window, 'innerHeight', {
  writable: true,
  configurable: true,
  value: 900,
});

// Provide requestAnimationFrame and cancelAnimationFrame fallbacks
if (!window.requestAnimationFrame) {
  window.requestAnimationFrame = (callback) => {
    return window.setTimeout(callback, 0);
  };
}

if (!window.cancelAnimationFrame) {
  window.cancelAnimationFrame = (id) => {
    window.clearTimeout(id);
  };
}

// Clear localStorage after each test
afterEach(() => {
  window.localStorage.clear();
});

// Ensure TextEncoder/TextDecoder exist
if (typeof global.TextEncoder === 'undefined') {
  const util = require('util');
  global.TextEncoder = util.TextEncoder;
  global.TextDecoder = util.TextDecoder;
}
