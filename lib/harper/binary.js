import { BinaryModuleImpl } from "./BinaryModule-S5UABkWQ.js";
const binaryUrl = "" + new URL("harper_wasm_bg.wasm", import.meta.url).href;
const binary = /* @__PURE__ */ BinaryModuleImpl.create(binaryUrl, "full");
export {
  binary
};
