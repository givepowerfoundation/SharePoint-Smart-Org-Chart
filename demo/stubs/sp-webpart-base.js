'use strict';
// Minimal stub so "extends BaseClientSideWebPart" compiles and runs without SPFx.
function BaseClientSideWebPart() {}
BaseClientSideWebPart.prototype = Object.create(null);
BaseClientSideWebPart.prototype.constructor = BaseClientSideWebPart;
module.exports = { BaseClientSideWebPart };
