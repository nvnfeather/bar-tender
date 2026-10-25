/* Worker bridge. No measurement data is sent to a server. */
window.MatplotlibEngine = class {
  constructor(onStatus) {
    this.onStatus = onStatus;
    this.nextId = 0;
    this.pending = new Map();
    this.worker = new Worker(new URL('matplotlib-worker.js', document.baseURI));
    this.worker.onmessage = ({data}) => {
      if (data.type === 'status') { this.onStatus(data.message); return; }
      const request = this.pending.get(data.id);
      if (!request) return;
      this.pending.delete(data.id);
      if (data.type === 'error') request.reject(new Error(data.message));
      else request.resolve(data);
    };
    this.worker.onerror = () => {
      this.failed = new Error('The chart engine could not start. Check your internet connection and retry.');
      for (const request of this.pending.values()) request.reject(new Error('The chart engine could not start. Check your internet connection and retry.'));
      this.pending.clear();
    };
  }
  warmup() { return this.request({type: 'warmup'}); }
  render(payload) { return this.request({type: 'render', payload}); }
  request(message) {
    return new Promise((resolve,reject) => {
      if (this.failed) { reject(this.failed); return; }
      const id = ++this.nextId;
      this.pending.set(id,{resolve,reject});
      this.worker.postMessage({id,...message});
    });
  }
  destroy() {
    this.worker.terminate();
    for (const request of this.pending.values()) request.reject(new Error('Chart engine restarted.'));
    this.pending.clear();
  }
};
