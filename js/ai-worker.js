/* XD3 PDF Editor — offline AI writing assistant (a module worker). Qwen2.5-0.5B-Instruct (Apache 2.0) runs on this device with
   transformers.js + ONNX Runtime (WebAssembly). Everything is loaded from the app's own lib/ai folder; nothing goes to the internet.
   Messages in:  { id, op: 'load' } · { id, op: 'run', messages, max, temp } · { id, op: 'stop' }
   Messages out: { id, type: 'progress', file, loaded, total } · { id, type: 'ready' } · { id, type: 'token', text } · { id, type: 'done', text } · { id, type: 'error', message } */
import { env, AutoTokenizer, AutoModelForCausalLM, TextStreamer, InterruptableStoppingCriteria } from '../lib/ai/transformers.min.js';

const MODEL = 'onnx-community/Qwen2.5-0.5B-Instruct';
env.allowRemoteModels = false; env.allowLocalModels = true;
env.localModelPath = new URL('../lib/ai/models/', import.meta.url).href;
env.useBrowserCache = false; // (the app's own offline cache already keeps lib/ — no second copy of 500 MB)
env.backends.onnx.wasm.wasmPaths = new URL('../lib/ai/', import.meta.url).href;
env.backends.onnx.wasm.numThreads = self.crossOriginIsolated ? Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 2) - 1)) : 1;

let tok = null, model = null, loading = null;
const stopper = new InterruptableStoppingCriteria();
function load(id) {
  return loading || (loading = (async () => {
    const progress_callback = p => { if (p.status === 'progress' && /\.onnx$|tokenizer\.json$/.test(p.file || '')) postMessage({ id, type: 'progress', file: p.file, loaded: p.loaded, total: p.total }); };
    tok = await AutoTokenizer.from_pretrained(MODEL, { progress_callback });
    model = await AutoModelForCausalLM.from_pretrained(MODEL, { dtype: self.__dtype || 'q4f16', device: 'wasm', progress_callback });
  })().catch(e => { loading = null; throw e; }));
}
self.onmessage = async ({ data: m }) => {
  const { id } = m; if (m.dtype) self.__dtype = m.dtype;
  try {
    if (m.op === 'stop') { stopper.interrupt(); return; }
    await load(id);
    if (m.op === 'load') return postMessage({ id, type: 'ready', threads: env.backends.onnx.wasm.numThreads });
    if (m.op === 'run') {
      stopper.reset();
      const inputs = tok.apply_chat_template(m.messages, { add_generation_prompt: true, return_dict: true });
      let text = '';
      const streamer = new TextStreamer(tok, { skip_prompt: true, skip_special_tokens: true, callback_function: t => { text += t; postMessage({ id, type: 'token', text: t }); } });
      await model.generate({ ...inputs, max_new_tokens: m.max || 400, do_sample: (m.temp || 0) > 0, temperature: m.temp || undefined, top_p: m.temp ? 0.9 : undefined, repetition_penalty: m.rep || 1.0, streamer, stopping_criteria: stopper });
      postMessage({ id, type: 'done', text });
    }
  } catch (e) { postMessage({ id, type: 'error', message: String(e && e.message || e) }); }
};
