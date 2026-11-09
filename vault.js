/* 허브·어드민 공용: 비밀번호 잠금 + 고료 링크 목록 복호화
   vault.json 은 암호문이라 공개돼도 코드·금액을 알 수 없어요. */
(function () {
  const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
  let vaultCache = null;

  async function getVault() {
    if (!vaultCache) {
      const r = await fetch('vault.json?t=' + Date.now(), { cache: 'no-store' });
      vaultCache = await r.json();
    }
    return vaultCache;
  }

  /* 맞는 비밀번호면 고료 링크 목록을, 틀리면 null */
  async function openVault(pin) {
    try {
      const v = await getVault();
      const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveKey']);
      const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: unb64(v.salt), iterations: v.iterations, hash: 'SHA-256' },
        base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
      const pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(v.iv) }, key, unb64(v.ct));
      return JSON.parse(new TextDecoder().decode(pt)).tiers;
    } catch { return null; }
  }

  async function sha256hex(s) {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
    return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
  }
  /* 공개 파일(brands.json)에는 고료 코드 대신 이 해시만 저장 */
  const tierHash = (code) => sha256hex('siriai-tier:' + code);

  const KEY = 'siriai-hub-pin';

  /* 4칸 PIN 입력 화면 연결. 성공하면 onOpen(tiers) */
  function mountLock({ lockEl, inputs, errEl, form, onOpen }) {
    let fails = 0, waitUntil = 0, busy = false;
    async function check() {
      if (busy || Date.now() < waitUntil) return;
      busy = true; errEl.textContent = '확인 중…';
      const pin = inputs.map((x) => x.value).join('');
      const tiers = await openVault(pin);
      busy = false;
      if (tiers) {
        try { sessionStorage.setItem(KEY, pin); } catch {}
        errEl.textContent = ''; lockEl.hidden = true; onOpen(tiers);
      } else {
        fails++;
        const box = inputs[0].parentElement; box.classList.remove('shake'); void box.offsetWidth; box.classList.add('shake');
        inputs.forEach((x) => { x.value = ''; }); inputs[0].focus();
        if (fails >= 5) { waitUntil = Date.now() + 30000; errEl.textContent = '여러 번 틀렸어요. 30초 후 다시 시도해주세요.'; fails = 0; }
        else errEl.textContent = '비밀번호가 맞지 않아요.';
      }
    }
    inputs.forEach((el, i) => {
      el.addEventListener('input', () => {
        el.value = el.value.replace(/\D/g, '').slice(-1);
        if (el.value && i < inputs.length - 1) inputs[i + 1].focus();
        if (inputs.every((x) => x.value)) check();
      });
      el.addEventListener('keydown', (e) => { if (e.key === 'Backspace' && !el.value && i > 0) inputs[i - 1].focus(); });
      el.addEventListener('paste', (e) => {
        const d = (e.clipboardData.getData('text') || '').replace(/\D/g, '').slice(0, inputs.length);
        if (!d) return; e.preventDefault();
        d.split('').forEach((c, j) => { inputs[j].value = c; });
        if (d.length === inputs.length) check(); else inputs[d.length].focus();
      });
    });
    form.addEventListener('submit', (e) => { e.preventDefault(); check(); });

    (async () => {
      let saved = null;
      try { saved = sessionStorage.getItem(KEY); } catch {}
      const tiers = saved ? await openVault(saved) : null;
      if (tiers) { lockEl.hidden = true; onOpen(tiers); } else inputs[0].focus();
    })();
  }

  function lock() { try { sessionStorage.removeItem(KEY); } catch {} location.reload(); }

  /* 캠페인별 고료처럼 공개 파일에 넣되 PIN 없이는 못 읽는 짧은 글 (형식: 'enc1:iv.ct') */
  const b64 = (u8) => btoa(String.fromCharCode(...u8));
  let fieldKeyP = null;
  function fieldKey() {
    if (!fieldKeyP) fieldKeyP = (async () => {
      const pin = sessionStorage.getItem(KEY); if (!pin) throw new Error('locked');
      const v = await getVault();
      const salt = new Uint8Array([...unb64(v.salt), ...new TextEncoder().encode('siriai-field')]);
      const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveKey']);
      return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: v.iterations, hash: 'SHA-256' },
        base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
    })().catch((e) => { fieldKeyP = null; throw e; });
    return fieldKeyP;
  }
  /* enc2: 구글 시트(Apps Script)도 풀 수 있는 방식 — HMAC-SHA256 키스트림 + 위변조 확인 태그.
     키(32바이트)는 PIN에서 만들고, 같은 키를 Apps Script 코드(비공개)에 넣어둬요. */
  let sheetKeyP = null;
  function sheetKey() {
    if (!sheetKeyP) sheetKeyP = (async () => {
      const pin = sessionStorage.getItem(KEY); if (!pin) throw new Error('locked');
      const v = await getVault();
      const salt = new Uint8Array([...unb64(v.salt), ...new TextEncoder().encode('siriai-sheet')]);
      const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveBits']);
      const raw = new Uint8Array(await crypto.subtle.deriveBits({ name: 'PBKDF2', salt, iterations: v.iterations, hash: 'SHA-256' }, base, 256));
      return { raw, key: await crypto.subtle.importKey('raw', raw, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']) };
    })().catch((e) => { sheetKeyP = null; throw e; });
    return sheetKeyP;
  }
  const cat = (...parts) => { const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
  const txt = (s) => new TextEncoder().encode(s);
  async function hmac(key, data) { return new Uint8Array(await crypto.subtle.sign('HMAC', key, data)); }
  async function xorStream(key, nonce, data) {
    const out = new Uint8Array(data.length);
    for (let i = 0; i * 32 < data.length; i++) {
      const ks = await hmac(key, cat(nonce, txt('ks'), new Uint8Array([i])));
      for (let j = 0; j < 32 && i * 32 + j < data.length; j++) out[i * 32 + j] = data[i * 32 + j] ^ ks[j];
    }
    return out;
  }
  async function seal(text) {
    if (!text) return '';
    const { key } = await sheetKey();
    const nonce = crypto.getRandomValues(new Uint8Array(16));
    const ct = await xorStream(key, nonce, txt(text).slice(0, 2000));
    const tag = (await hmac(key, cat(txt('tag'), nonce, ct))).slice(0, 16);
    return 'enc2:' + b64(nonce) + '.' + b64(ct) + '.' + b64(tag);
  }
  async function unseal(s) {
    s = String(s || '');
    try {
      if (s.startsWith('enc2:')) {
        const { key } = await sheetKey();
        const [nonce, ct, tag] = s.slice(5).split('.').map(unb64);
        const want = (await hmac(key, cat(txt('tag'), nonce, ct))).slice(0, 16);
        if (want.some((x, i) => x !== tag[i])) return '';
        return new TextDecoder().decode(await xorStream(key, nonce, ct));
      }
      if (s.startsWith('enc1:')) {
        const [iv, ct] = s.slice(5).split('.');
        return new TextDecoder().decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(iv) }, await fieldKey(), unb64(ct)));
      }
    } catch {}
    return '';
  }
  /* enc3: 링크 코드로 잠그는 고료 — 그 링크로 들어온 가이드만 풀 수 있어요 (index.html openFee 와 짝) */
  async function codeKey(code, use) {
    const raw = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('siriai-fee-key:' + code));
    return crypto.subtle.importKey('raw', raw, 'AES-GCM', false, [use]);
  }
  async function sealForCode(code, text) {
    if (!text) return '';
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await codeKey(code, 'encrypt'), new TextEncoder().encode(text)));
    return 'enc3:' + b64(iv) + '.' + b64(ct);
  }
  async function openForCode(code, s) {
    try {
      const [iv, ct] = String(s).replace(/^enc3:/, '').split('.').map(unb64);
      return new TextDecoder().decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, await codeKey(code, 'decrypt'), ct));
    } catch { return ''; }
  }

  /* Apps Script 에 넣을 키 (어드민 '지원서 접수' 카드에서만 보여줌) */
  async function sheetKeyB64() { return b64((await sheetKey()).raw); }

  window.SiriaiVault = { openVault, tierHash, mountLock, lock, seal, unseal, sheetKeyB64, sealForCode, openForCode };
})();
