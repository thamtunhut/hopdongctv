// Giai ma base64 (tu embeddedAssets.js) thanh ArrayBuffer, dung thay cho fetch()
// vi fetch() bi trinh duyet chan khi mo file bang double-click (file://).
(function (global) {
  function base64ToArrayBuffer(b64) {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }

  function getEmbeddedAsset(name) {
    const b64 = global.EMBEDDED_ASSETS_B64 && global.EMBEDDED_ASSETS_B64[name];
    if (!b64) throw new Error(`Khong tim thay asset nhung san: ${name}`);
    return base64ToArrayBuffer(b64);
  }

  global.getEmbeddedAsset = getEmbeddedAsset;
})(window);
