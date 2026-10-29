'use strict';

(() => {
  const token = location.hash.slice(1);
  const screen = document.getElementById('screen');
  const typing = document.getElementById('typing');
  const status = document.getElementById('status');
  const errorBox = document.getElementById('error');
  let current;
  let stopped = false;
  let refreshing = false;
  let actions = Promise.resolve();
  let polling;

  function error(message) {
    errorBox.textContent = message;
    errorBox.hidden = false;
  }

  function finish(state) {
    stopped = true;
    clearInterval(polling);
    typing.value = '';
    document.getElementById('controls').hidden = true;
    document.getElementById('screen-container').hidden = true;
    screen.removeAttribute('src');
    if (state === 'complete') {
      document.getElementById('done').hidden = false;
      status.textContent = 'Microsoft sign-in complete';
      errorBox.hidden = true;
    } else status.textContent = 'Login link closed. Call the MCP login tool to continue with a new link.';
  }

  async function api(route, body) {
    const response = await fetch(route, {
      method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${token}`, ...(body === undefined ? {} : { 'Content-Type': 'application/json' }) },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      cache: 'no-store', referrerPolicy: 'no-referrer', signal: AbortSignal.timeout(20000),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Could not reach the login session.');
    return data;
  }

  async function refresh() {
    if (stopped || refreshing) return;
    refreshing = true;
    try {
      const state = await api('/api/state');
      if (['complete', 'cancelled', 'expired'].includes(state.state)) return finish(state.state);
      if (state.state !== 'pending') { status.textContent = 'The sign-in screen is loading…'; return; }
      current = state;
      screen.src = `data:image/png;base64,${state.image}`;
      screen.hidden = false;
      document.getElementById('origin').textContent = state.origin;
      status.textContent = 'Connected · live sign-in screen';
      typing.disabled = false;
      document.querySelectorAll('[data-key]').forEach(button => { button.disabled = false; });
    } catch (failure) {
      if (!stopped) error(`${failure.message} If the link expired, call login again in your MCP client.`);
    } finally { refreshing = false; }
  }

  function action(value, focus = false) {
    if (stopped || !current) return;
    const viewId = current.viewId;
    actions = actions.then(async () => {
      if (stopped) return;
      try {
        const result = await api('/api/action', { viewId, action: value });
        errorBox.hidden = true;
        if (result.state === 'complete') return finish('complete');
        if (focus) typing.focus({ preventScroll: true });
        await refresh();
      } catch (failure) {
        if (!stopped) error(failure.message);
        await refresh();
      }
    });
  }

  function text(event) {
    if (event.isComposing) return;
    const value = typing.value;
    typing.value = '';
    if (value) action({ type: 'type', text: value });
  }

  typing.addEventListener('input', text);
  typing.addEventListener('compositionend', text);
  typing.addEventListener('keydown', event => {
    let key = event.key;
    if ((event.ctrlKey || event.metaKey) && key.toLowerCase() === 'a') key = event.metaKey ? 'Meta+A' : 'Control+A';
    if (['Enter', 'Backspace', 'Delete', 'Escape', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'Control+A', 'Meta+A'].includes(key)) {
      event.preventDefault();
      action({ type: 'key', key });
    }
  });
  screen.addEventListener('click', event => {
    const bounds = screen.getBoundingClientRect();
    action({ type: 'click', x: (event.clientX - bounds.left) * current.width / bounds.width,
      y: (event.clientY - bounds.top) * current.height / bounds.height }, true);
  });
  screen.addEventListener('wheel', event => {
    event.preventDefault();
    action({ type: 'scroll', deltaY: Math.max(-3000, Math.min(3000, event.deltaY)) });
  }, { passive: false });
  document.querySelectorAll('[data-key]').forEach(button => button.addEventListener('click', () => action({ type: 'key', key: button.dataset.key }, true)));
  document.getElementById('refresh').addEventListener('click', refresh);
  document.getElementById('cancel').addEventListener('click', async () => {
    try { await api('/api/cancel', {}); finish('cancelled'); }
    catch (failure) { error(failure.message); }
  });
  if (!/^[A-Za-z0-9_-]{43}$/.test(token)) {
    finish('cancelled');
    error('Open the complete private link returned by the MCP login tool.');
  } else {
    refresh();
    polling = setInterval(() => { if (!document.hidden) refresh(); }, 1200);
  }
})();
