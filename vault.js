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

  window.SiriaiVault = { openVault, tierHash, mountLock, lock };
})();
