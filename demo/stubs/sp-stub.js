'use strict';
// Generic stub for SPFx packages not needed at runtime in demo mode.
// Returns a function (usable as a constructor/class) for any named export.
function noop() {}
module.exports = new Proxy({ __esModule: true }, {
  get: function(_, key) { return key === '__esModule' ? true : noop; },
});
