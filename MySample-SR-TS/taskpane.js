!function(){"use strict";var e,t,n,o,r={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},98362:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},58394:function(e,t,n){e.exports=n.p+"1fda685b81e1123773f6.css"}},i={};function a(e){var t=i[e];if(void 0!==t)return t.exports;var n=i[e]={exports:{}};return r[e](n,n.exports,a),n.exports}a.m=r,a.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return a.d(t,{a:t}),t},a.d=function(e,t){for(var n in t)a.o(t,n)&&!a.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},a.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),a.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;a.g.importScripts&&(e=a.g.location+"");var t=a.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),a.p=e}(),a.b=document.baseURI||self.location.href,function(){(async()=>{await Office.onReady(),console.log("Office is ready"),document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=r,await r()})();const e="MYTSSAMPLE",t="PRODUCTLINKEDENTITYSERVICE",n={dataProvider:e,id:"categories",name:"Cateogories",loadFunctionId:t},o={dataProvider:e,id:"suppliers",name:"Suppliers",loadFunctionId:t};async function r(){const r={dataProvider:e,id:"products",name:"Products",loadFunctionId:t,periodicRefreshInterval:300,supportedRefreshModes:["Periodic","OnLoad"]};await Excel.run((async e=>{const t=e.workbook.linkedEntityDataDomains;t.add(r),t.add(n),t.add(o),await e.sync(),console.log("Linked entity data domains registered.")}))}}(),e=a(14385),t=a.n(e),n=new URL(a(58394),a.b),o=new URL(a(98362),a.b),t()(n),t()(o)}();
//# sourceMappingURL=taskpane.js.map