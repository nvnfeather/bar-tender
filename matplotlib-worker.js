/* Python and Matplotlib run here, away from the editor's UI thread. */
const PYODIDE_URL = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let runtime;
let bootPromise;
async function boot() {
  if (!bootPromise) bootPromise = (async () => {
    self.postMessage({type:'status', message:'Loading Python… First load may take a minute.'});
    importScripts(PYODIDE_URL + 'pyodide.js');
    runtime = await loadPyodide({indexURL:PYODIDE_URL});
    self.postMessage({type:'status', message:'Loading Matplotlib… Your measurements stay in this browser.'});
    const [response] = await Promise.all([fetch(new URL('plot_renderer.py',self.location.href)), runtime.loadPackage('matplotlib')]);
    if (!response.ok) throw new Error('Could not load the chart renderer.');
    await runtime.runPythonAsync(await response.text());
    self.postMessage({type:'status', message:'Matplotlib ready'});
  })();
  return bootPromise;
}
self.onmessage = async ({data}) => {
  try {
    await boot();
    if (data.type === 'warmup') {
      self.postMessage({type:'result', id:data.id, ready:true});
      return;
    }
    runtime.globals.set('plot_payload', JSON.stringify(data.payload));
    const result = await runtime.runPythonAsync('render_charts(plot_payload)');
    self.postMessage({type:'result', id:data.id, ...JSON.parse(result)});
  } catch (error) {
    self.postMessage({type:'error', id:data.id, message:String(error.message||error)});
  } finally {
    if (runtime && data.type !== 'warmup') runtime.globals.delete('plot_payload');
  }
};
