/* ============================================================
   utils.js — Funções utilitárias globais
   ============================================================ */

// ── Toast Notifications ──────────────────────────────────────
function showToast(msg, type = 'info', duration = 3500) {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.textContent = msg;
  container.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transform = 'translateX(20px)'; t.style.transition = '0.3s'; setTimeout(() => t.remove(), 300); }, duration);
}

// ── Formatação ────────────────────────────────────────────────
function formatBRL(val) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val || 0);
}

function formatNumPad(n) {
  return n.toString().padStart(2, '0');
}

function formatDate(ts) {
  if (!ts) return '—';
  const d = ts.toDate ? ts.toDate() : new Date(ts);
  return d.toLocaleDateString('pt-BR');
}

function formatDateTime(ts) {
  if (!ts) return '—';
  const d = ts.toDate ? ts.toDate() : new Date(ts);
  return d.toLocaleString('pt-BR', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });
}

// ── Número formatado para tags HTML ─────────────────────────
function numTagsHtml(nums) {
  return (nums || []).map(n => `<span class="num-tag">${formatNumPad(n)}</span>`).join(' ');
}

// ── Modal helpers ─────────────────────────────────────────────
function openModal(id)  { document.getElementById(id)?.classList.add('open'); }
function closeModal(id) { document.getElementById(id)?.classList.remove('open'); }
// Fechar ao clicar fora
document.addEventListener('click', e => {
  if (e.target.classList.contains('modal-overlay')) {
    e.target.classList.remove('open');
  }
});

// ── Validação simples ─────────────────────────────────────────
function validatePhone(tel) {
  return /^\(?\d{2}\)?\s?\d{4,5}-?\d{4}$/.test(tel.replace(/\s/g, ''));
}

// ── Gerar QR Code PIX (payload estático simplificado) ─────────
// Em produção, use a API do seu banco para PIX dinâmico
function buildPixPayload(chave, nome, cidade, valor) {
  const v = valor.toFixed(2);
  const gui  = 'BR.GOV.BCB.PIX';
  const pixField = tlv('00', 'BR.GOV.BCB.PIX') + tlv('01', chave);
  const merchantInfo = tlv('26', tlv('00', gui) + tlv('01', chave));
  const valorStr = tlv('54', v);
  const payload = [
    tlv('00', '01'),
    merchantInfo,
    tlv('52', '0000'),
    tlv('53', '986'),
    valorStr,
    tlv('58', 'BR'),
    tlv('59', nome.substring(0, 25)),
    tlv('60', cidade.substring(0, 15)),
    tlv('62', tlv('05', '***')),
  ].join('');
  return payload + tlv('63', crc16(payload + '6304'));
}
function tlv(tag, val) {
  const l = val.length.toString().padStart(2, '0');
  return tag + l + val;
}
function crc16(str) {
  let crc = 0xFFFF;
  for (let i = 0; i < str.length; i++) {
    crc ^= str.charCodeAt(i) << 8;
    for (let j = 0; j < 8; j++) crc = (crc & 0x8000) ? ((crc << 1) ^ 0x1021) : (crc << 1);
  }
  return ((crc & 0xFFFF).toString(16).toUpperCase().padStart(4, '0'));
}

// ── Copiar para clipboard ─────────────────────────────────────
function copyText(text, btnEl) {
  navigator.clipboard.writeText(text).then(() => {
    if (btnEl) { const orig = btnEl.textContent; btnEl.textContent = 'Copiado!'; setTimeout(() => btnEl.textContent = orig, 1800); }
    showToast('Copiado para a área de transferência!', 'success');
  });
}

// ── Admin Tab switcher ────────────────────────────────────────
function switchAdminTab(tabName) {
  document.querySelectorAll('.admin-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tabName));
  document.querySelectorAll('.admin-content').forEach(c => c.classList.toggle('active', c.id === `tab-${tabName}`));
}

// ── Obter parâmetro de URL ────────────────────────────────────
function getParam(name) {
  return new URLSearchParams(window.location.search).get(name);
}
