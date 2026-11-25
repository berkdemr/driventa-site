// Driventa İşletme – Panel arayüzü
import { api } from "./api.js";
import { readXlsx, readCsv, writeXlsx, download } from "./xlsx.js";

const $ = s => document.querySelector(s);
const root = $("#root");
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const I = {
  home: '<svg viewBox="0 0 24 24"><path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/></svg>',
  car: '<svg viewBox="0 0 24 24"><path d="M5 17h14M3 13l2-5.5A2 2 0 0 1 6.9 6h10.2a2 2 0 0 1 1.9 1.5L21 13v4a1 1 0 0 1-1 1h-1.5a1 1 0 0 1-1-1v-1h-11v1a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/><path d="M3 13h18"/></svg>',
  check: '<svg viewBox="0 0 24 24"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V3h6v1M9 13l2 2 4-4"/></svg>',
  users: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M21.5 20a6.5 6.5 0 0 0-4-6"/></svg>',
  wrench: '<svg viewBox="0 0 24 24"><path d="M14.5 6.5a4 4 0 0 1 5.4-3.7l-2.6 2.6.8 2 2 .8 2.6-2.6a4 4 0 0 1-5.4 5.4L8.5 19.8a2 2 0 0 1-2.8-2.8l8.8-8.8a4 4 0 0 1 0-1.7z"/></svg>',
  cog: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
  out: '<svg viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/></svg>',
  menu: '<svg viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/></svg>',
  ok: '<svg viewBox="0 0 24 24"><path d="m5 12 5 5 9-10"/></svg>',
  copy: '<svg viewBox="0 0 24 24"><rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>',
  up: '<svg viewBox="0 0 24 24"><path d="M12 16V4M7 9l5-5 5 5M5 20h14"/></svg>',
  down: '<svg viewBox="0 0 24 24"><path d="M12 4v12M7 11l5 5 5-5M5 20h14"/></svg>',
  edit: '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
  x: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  search: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
  refresh: '<svg viewBox="0 0 24 24"><path d="M21 12a9 9 0 1 1-3-6.7L21 8M21 3v5h-5"/></svg>'
};
const G_LOGO = '<svg viewBox="0 0 48 48" style="stroke:none"><path fill="#FFC107" d="M43.6 20.1H42V20H24v8h11.3C33.7 32.7 29.2 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 13 4 4 13 4 24s9 20 20 20 20-9 20-20c0-1.3-.1-2.6-.4-3.9z"/><path fill="#FF3D00" d="m6.3 14.7 6.6 4.8C14.7 15.1 19 12 24 12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2C29.2 35.1 26.7 36 24 36c-5.2 0-9.6-3.3-11.3-8l-6.5 5C9.5 39.6 16.2 44 24 44z"/><path fill="#1976D2" d="M43.6 20.1H42V20H24v8h11.3c-.8 2.2-2.2 4.2-4.1 5.6l6.2 5.2C37 39.2 44 34 44 24c0-1.3-.1-2.6-.4-3.9z"/></svg>';

/* ---------- Hata mesajları (Türkçe) ---------- */
function errMsg(e) {
  const c = (e && (e.code || e.message)) || "";
  const map = {
    "auth/invalid-credential": "E-posta veya şifre hatalı.",
    "auth/wrong-password": "E-posta veya şifre hatalı.",
    "auth/user-not-found": "Bu e-postayla kayıtlı bir hesap bulunamadı.",
    "auth/invalid-email": "Geçerli bir e-posta adresi girin.",
    "auth/email-already-in-use": "Bu e-posta adresiyle zaten bir hesap var. Giriş yapmayı deneyin.",
    "auth/weak-password": "Şifre en az 6 karakter olmalı.",
    "auth/too-many-requests": "Çok fazla deneme yapıldı. Lütfen birkaç dakika sonra tekrar deneyin.",
    "auth/network-request-failed": "İnternet bağlantısı kurulamadı.",
    "auth/popup-closed-by-user": "Google penceresi kapatıldı.",
    "auth/popup-blocked": "Tarayıcı açılır pencereyi engelledi. Lütfen izin verip tekrar deneyin.",
    "auth/unauthorized-domain": "Bu adresten giriş yapılmasına izin verilmiyor (Firebase yetkili alan adları).",
    "permission-denied": "Bu işlem için yetkin yok.",
    "unavailable": "Sunucuya ulaşılamıyor. İnternet bağlantını kontrol et."
  };
  for (const k in map) if (c.includes(k)) return map[k];
  console.error(e);
  return "Beklenmeyen bir hata oluştu. Lütfen tekrar deneyin.";
}
let toastT;
function toast(msg, err) {
  const t = $("#toast"); t.textContent = msg; t.classList.toggle("err", !!err); t.classList.add("show");
  clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), 2600);
}
async function busy(btn, fn) {
  const txt = btn.innerHTML; btn.disabled = true; btn.innerHTML = '<span class="spinner" style="width:18px;height:18px;border-width:2px;margin:0"></span>';
  try { return await fn(); } finally { btn.disabled = false; btn.innerHTML = txt; }
}

/* ---------- Durum ---------- */
const S = { user: null, profile: null, company: null, member: null, page: "overview", members: null, vehicles: null };

/* ---------- Giriş ekranı ---------- */
function renderAuth(mode = "login") {
  root.innerHTML = `
  <div class="auth">
    <div class="auth-side">
      <div class="brand"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      <div>
        <h2>Filonu <span>tek panelden</span> yönet.</h2>
        <p class="muted">Şoförlerin masrafı telefondan girer, sen onaylar ve tüm filonun maliyetini anlık görürsün.</p>
        <ul>
          <li>${I.ok}Şoför kayıtları yönetici onayıyla</li>
          <li>${I.ok}Araç bazlı maliyet ve km başı gider</li>
          <li>${I.ok}Bakım, sigorta ve muayene takvimi</li>
          <li>${I.ok}Veriler Avrupa'daki sunucularda</li>
        </ul>
      </div>
      <p class="muted" style="font-size:13px">© ${new Date().getFullYear()} BeSey · <a class="link" href="https://driventa.com.tr/gizlilik.html">Gizlilik</a></p>
    </div>
    <div class="auth-main"><div class="auth-card">
      <div class="brand" style="margin-bottom:22px"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      <h1>${mode === "reset" ? "Şifreni sıfırla" : "Hoş geldin"}</h1>
      <p class="muted">${mode === "reset" ? "E-postana bir sıfırlama bağlantısı gönderelim." : "Yönetici paneline giriş yap veya şirketin için hesap oluştur."}</p>
      ${mode === "reset" ? "" : `<div class="tabs"><button data-mode="login" class="${mode === "login" ? "on" : ""}">Giriş yap</button><button data-mode="register" class="${mode === "register" ? "on" : ""}">Hesap oluştur</button></div>`}
      <div id="authMsg"></div>
      <form id="authForm" novalidate style="${mode === "reset" ? "margin-top:22px" : ""}">
        ${mode === "register" ? `<div class="field"><label>Ad soyad</label><input class="input" name="name" autocomplete="name" maxlength="60"></div>` : ""}
        <div class="field"><label>E-posta</label><input class="input" name="email" type="email" autocomplete="email"></div>
        ${mode === "reset" ? "" : `<div class="field"><label>Şifre</label><input class="input" name="pw" type="password" autocomplete="${mode === "register" ? "new-password" : "current-password"}"></div>`}
        <button class="btn btn-primary btn-block" type="submit">${mode === "login" ? "Giriş yap" : mode === "register" ? "Hesap oluştur" : "Bağlantı gönder"}</button>
      </form>
      ${mode === "reset" ? `<p style="margin-top:16px;text-align:center"><a class="link" href="#" data-mode="login">← Girişe dön</a></p>` : `
      <div class="or">veya</div>
      <button class="gbtn" id="gBtn">${G_LOGO} Google ile devam et</button>
      ${mode === "login" ? `<p style="margin-top:16px;text-align:center"><a class="link" href="#" data-mode="reset">Şifremi unuttum</a></p>` : ""}`}
      <p class="muted" style="font-size:12.5px;margin-top:22px;text-align:center">Şoför müsün? Driventa uygulamasını telefonuna kur ve yöneticinin verdiği katılım koduyla şirkete katıl.</p>
    </div></div>
  </div>`;
  root.querySelectorAll("[data-mode]").forEach(b => b.addEventListener("click", e => { e.preventDefault(); renderAuth(b.dataset.mode); }));
  const msg = (t, ok) => { $("#authMsg").innerHTML = `<div class="alert ${ok ? "ok" : "err"}">${esc(t)}</div>`; };
  const f = $("#authForm");
  f.addEventListener("submit", e => {
    e.preventDefault();
    const email = f.email.value.trim(), pw = f.pw ? f.pw.value : "", name = f.name ? f.name.value.trim() : "";
    if (mode === "register" && name.length < 2) return msg("Ad soyad girin.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return msg("Geçerli bir e-posta adresi girin.");
    if (mode !== "reset" && pw.length < 6) return msg("Şifre en az 6 karakter olmalı.");
    busy(f.querySelector("button[type=submit]"), async () => {
      try {
        if (mode === "login") await api.signIn(email, pw);
        else if (mode === "register") await api.signUp(name, email, pw);
        else { await api.resetPassword(email); msg("Sıfırlama bağlantısı e-postana gönderildi. Gelen kutunu (ve spam klasörünü) kontrol et.", true); }
      } catch (err) { msg(errMsg(err)); }
    });
  });
  const g = $("#gBtn");
  if (g) g.addEventListener("click", () => busy(g, async () => { try { await api.signInGoogle(); } catch (err) { msg(errMsg(err)); } }));
}

/* ---------- Şirket oluşturma ---------- */
function renderOnboarding() {
  root.innerHTML = `
  <div class="center"><div class="card">
    <div class="brand" style="margin-bottom:20px"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
    <h1 style="font-size:24px">Şirketini oluştur</h1>
    <p class="muted" style="margin:6px 0 20px">Bu bilgilerle filonun yönetim paneli oluşturulacak. Daha sonra Ayarlar'dan değiştirebilirsin.</p>
    <div id="obMsg"></div>
    <form id="obForm" novalidate>
      <div class="field"><label>Şirket adı *</label><input class="input" name="name" maxlength="80" placeholder="Örn. Demir Lojistik"></div>
      <div class="field"><label>Telefon</label><input class="input" name="phone" maxlength="20" inputmode="tel" placeholder="05xx xxx xx xx"></div>
      <div class="field"><label>Araç sayısı</label>
        <select class="input" name="size"><option value="">Seçin</option><option>1–5</option><option>6–20</option><option>21–50</option><option>51–200</option><option>200+</option></select></div>
      <button class="btn btn-primary btn-block" type="submit">Şirketi oluştur</button>
    </form>
    <p style="margin-top:16px;text-align:center"><a class="link" href="#" id="obOut">Farklı hesapla giriş yap</a></p>
  </div></div>`;
  $("#obOut").addEventListener("click", e => { e.preventDefault(); api.signOut(); });
  const f = $("#obForm");
  f.addEventListener("submit", e => {
    e.preventDefault();
    const name = f.name.value.trim();
    if (name.length < 2) { $("#obMsg").innerHTML = '<div class="alert err">Şirket adını girin.</div>'; return; }
    busy(f.querySelector("button"), async () => {
      try {
        const cid = await api.createCompany(S.user, { name, phone: f.phone.value.trim(), size: f.size.value });
        await loadCompany(cid);
        toast("Şirketin oluşturuldu. Hoş geldin!");
      } catch (err) { $("#obMsg").innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; }
    });
  });
}

/* ---------- Şoför hesabı panele girerse ---------- */
function renderDriverInfo() {
  root.innerHTML = `
  <div class="center"><div class="card" style="text-align:center">
    <div class="empty" style="padding:10px 0 0">
      <div class="ic">${I.car}</div>
      <h3>Bu hesap bir şoför hesabı</h3>
      <p>Şoförler kayıtlarını Driventa mobil uygulamasından girer. Yönetim paneline yalnızca şirket sahibi ve yöneticiler erişebilir.</p>
    </div>
    <button class="btn btn-ghost" style="margin-top:20px" id="dOut">${I.out} Çıkış yap</button>
  </div></div>`;
  $("#dOut").addEventListener("click", () => api.signOut());
}

/* ---------- Panel iskeleti ---------- */
const PAGES = [
  { id: "overview", ad: "Özet", ic: "home" },
  { id: "vehicles", ad: "Araçlar", ic: "car" },
  { id: "approvals", ad: "Kayıt Onayları", ic: "check", soon: true },
  { id: "drivers", ad: "Şoförler", ic: "users" },
  { id: "maintenance", ad: "Bakım Takvimi", ic: "wrench", soon: true },
  { id: "settings", ad: "Ayarlar", ic: "cog" }
];

function renderShell() {
  const u = S.user, c = S.company, m = S.member;
  const nm = m.name || u.displayName || u.email;
  root.innerHTML = `
  <div class="shell">
    <aside class="side" id="side">
      <div class="brand"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      ${PAGES.map(p => `<button class="nav-item ${S.page === p.id ? "on" : ""}" data-page="${p.id}">${I[p.ic]}${p.ad}${p.soon ? '<span class="soon">Yakında</span>' : ""}</button>`).join("")}
      <div class="side-foot">
        <div class="me"><div class="avatar">${esc((nm[0] || "?").toLocaleUpperCase("tr-TR"))}</div>
          <div><b>${esc(nm)}</b><small>${m.role === "owner" ? "Şirket sahibi" : "Yönetici"}</small></div></div>
        <button class="nav-item" id="outBtn">${I.out}Çıkış yap</button>
      </div>
    </aside>
    <main class="main">
      <div class="top">
        <div style="display:flex;gap:12px;align-items:center">
          <button class="menu-btn" id="menuBtn" aria-label="Menü">${I.menu}</button>
          <div><h1 id="pgTitle"></h1><p>${esc(c.name)}</p></div>
        </div>
      </div>
      <div id="page"></div>
    </main>
  </div>`;
  root.querySelectorAll("[data-page]").forEach(b => b.addEventListener("click", () => { S.page = b.dataset.page; $("#side").classList.remove("open"); renderShell(); }));
  $("#outBtn").addEventListener("click", () => api.signOut());
  $("#menuBtn").addEventListener("click", e => { e.stopPropagation(); $("#side").classList.toggle("open"); });
  document.addEventListener("click", e => { const s = $("#side"); if (s && s.classList.contains("open") && !s.contains(e.target)) s.classList.remove("open"); }, { once: true });
  const page = PAGES.find(p => p.id === S.page);
  $("#pgTitle").textContent = page.ad;
  (Pages[S.page] || Pages.soon)(page);
}

const emptyBox = (ic, t, d) => `<div class="card empty"><div class="ic">${I[ic]}</div><h3>${t}</h3><p>${d}</p></div>`;

const Pages = {
  overview() {
    $("#page").innerHTML = `
      <div class="kpis">
        <div class="kpi"><small>Araç</small><b id="kVehicles">–</b></div>
        <div class="kpi"><small>Şoför</small><b id="kDrivers">–</b></div>
        <div class="kpi"><small>Bu ay masraf</small><b>₺0</b></div>
        <div class="kpi"><small>Onay bekleyen</small><b>0</b></div>
      </div>
      <div class="card">
        <h3 style="font-size:17px;margin-bottom:4px">Başlarken</h3>
        <p class="muted">Filonu Driventa'ya taşımak için şu adımları izle:</p>
        <ol class="steps">
          <li>Araçlarını ekle: <a class="link" href="#" data-goto="vehicles">Araçlar</a> sayfasından tek tek veya Excel'den toplu olarak.</li>
          <li>Şoförlerine katılım kodunu ver: <b style="color:var(--accent);letter-spacing:2px">${esc(S.company.joinCode)}</b></li>
          <li>Şoförler Driventa uygulamasından bu kodla şirkete katılsın.</li>
          <li>Her şoföre aracını ata; şoförlerin girdiği kayıtları onayla.</li>
        </ol>
      </div>`;
    $("#page").querySelectorAll("[data-goto]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); S.page = a.dataset.goto; renderShell(); }));
    api.listVehicles(S.company.id).then(vs => { S.vehicles = vs; const el = $("#kVehicles"); if (el) el.textContent = vs.length; }).catch(() => {});
    api.listMembers(S.company.id).then(ms => { S.members = ms; const el = $("#kDrivers"); if (el) el.textContent = ms.filter(x => x.role === "driver").length; })
      .catch(() => { const el = $("#kDrivers"); if (el) el.textContent = "–"; });
  },

  drivers() {
    const c = S.company;
    $("#page").innerHTML = `
      <div class="grid2">
        <div class="card">
          <h3 style="font-size:17px">Katılım kodu</h3>
          <p class="muted" style="margin:4px 0 16px">Şoförler Driventa uygulamasında <b>Profil → İşletme hesabı</b> bölümüne bu kodu girerek şirketine katılır.</p>
          <div class="code" id="jc">${esc(c.joinCode)}</div>
          <div style="display:flex;gap:10px;margin-top:16px;flex-wrap:wrap">
            <button class="btn btn-ghost" id="cpBtn">${I.copy} Kopyala</button>
            <button class="btn btn-ghost" id="rgBtn">${I.refresh} Yeni kod oluştur</button>
          </div>
          <p class="muted" style="font-size:12.5px;margin-top:12px">Kod başkasının eline geçtiyse yeni kod oluştur. Eski kod hemen geçersiz olur; katılmış şoförler etkilenmez.</p>
        </div>
        <div class="card">
          <h3 style="font-size:17px">Nasıl çalışır?</h3>
          <ol class="steps">
            <li>Şoför Driventa'yı telefonuna kurar ve e-postasıyla hesap açar.</li>
            <li>Katılım kodunu girer; listede "araç atanmadı" olarak görünür.</li>
            <li>Sen <b>Araçlar</b> sayfasından şoföre aracını atarsın.</li>
            <li>Şoförün girdiği kayıtlar senin onayına düşer.</li>
          </ol>
        </div>
      </div>
      <div class="card" style="margin-top:16px">
        <h3 style="font-size:17px;margin-bottom:10px">Ekip</h3>
        <div id="mList"><div class="spinner"></div></div>
      </div>`;
    $("#cpBtn").addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(S.company.joinCode); toast("Kod kopyalandı."); } catch (e) { toast("Kopyalanamadı, kodu elle seçebilirsin.", true); }
    });
    $("#rgBtn").addEventListener("click", e => {
      if (!confirmInline(e.currentTarget, "Emin misin? Tekrar bas")) return;
      busy(e.currentTarget, async () => {
        try { const code = await api.regenerateJoinCode(c.id, c.joinCode); S.company.joinCode = code; $("#jc").textContent = code; toast("Yeni katılım kodu oluşturuldu."); }
        catch (err) { toast(errMsg(err), true); }
      });
    });
    Promise.all([api.listMembers(c.id), api.listVehicles(c.id).catch(() => [])]).then(([ms, vs]) => {
      S.members = ms; S.vehicles = vs;
      const roleAd = { owner: "Sahip", admin: "Yönetici", driver: "Şoför" };
      $("#mList").innerHTML = ms.length ? `<div style="overflow-x:auto"><table class="tbl"><thead><tr><th>Ad</th><th>E-posta</th><th>Rol</th><th>Araç</th></tr></thead><tbody>
        ${ms.sort((a, b) => (a.role === "owner" ? -1 : 1)).map(m => `<tr><td>${esc(m.name || "–")}</td><td>${esc(m.email || "–")}</td>
          <td><span class="role ${m.role}">${roleAd[m.role] || m.role}</span></td>
          <td class="muted">${m.role === "driver" ? (m.vehicleIds && m.vehicleIds.length ? m.vehicleIds.map(id => { const v = (S.vehicles || []).find(x => x.id === id); return v ? esc(v.plaka) : ""; }).filter(Boolean).join(", ") || m.vehicleIds.length + " araç" : "Atanmadı") : "–"}</td></tr>`).join("")}
        </tbody></table></div>` : emptyBox("users", "Henüz ekip üyesi yok", "Katılım kodunu şoförlerinle paylaş.");
    }).catch(err => { $("#mList").innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; });
  },

  settings() {
    const c = S.company;
    $("#page").innerHTML = `
      <div class="grid2">
        <div class="card">
          <h3 style="font-size:17px;margin-bottom:14px">Şirket bilgileri</h3>
          <form id="stForm" novalidate>
            <div class="field"><label>Şirket adı *</label><input class="input" name="name" maxlength="80" value="${esc(c.name)}"></div>
            <div class="field"><label>Telefon</label><input class="input" name="phone" maxlength="20" value="${esc(c.phone || "")}"></div>
            <div class="field"><label>Araç sayısı</label><select class="input" name="size">
              ${["", "1–5", "6–20", "21–50", "51–200", "200+"].map(o => `<option value="${o}" ${o === (c.size || "") ? "selected" : ""}>${o || "Seçin"}</option>`).join("")}</select></div>
            <button class="btn btn-primary" type="submit">Kaydet</button>
          </form>
        </div>
        <div class="card">
          <h3 style="font-size:17px;margin-bottom:10px">Hesap</h3>
          <p class="muted">Giriş yapılan hesap</p><p style="margin-bottom:14px">${esc(S.user.email)}</p>
          <p class="muted">Paket</p><p style="margin-bottom:14px">Deneme (erken erişim)</p>
          <p class="muted" style="font-size:13px">Verilerin Avrupa'daki (eur3) sunucularda saklanır. <a class="link" href="https://driventa.com.tr/gizlilik.html" target="_blank" rel="noopener">Gizlilik politikası</a></p>
        </div>
      </div>`;
    const f = $("#stForm");
    f.addEventListener("submit", e => {
      e.preventDefault();
      const name = f.name.value.trim();
      if (name.length < 2) return toast("Şirket adını girin.", true);
      busy(f.querySelector("button"), async () => {
        try {
          const data = { name, phone: f.phone.value.trim(), size: f.size.value };
          await api.updateCompany(c.id, data); Object.assign(S.company, data); toast("Kaydedildi."); renderShell();
        } catch (err) { toast(errMsg(err), true); }
      });
    });
  },

  soon(page) {
    const d = {
      vehicles: "Araç ekleme, düzenleme ve Excel'den toplu aktarım bir sonraki aşamada (F2) geliyor.",
      approvals: "Şoförlerin girdiği kayıtları onaylama ve reddetme ekranı F3 aşamasında geliyor.",
      maintenance: "Tüm filonun bakım, sigorta ve muayene takvimi F5 aşamasında geliyor."
    };
    $("#page").innerHTML = emptyBox(page.ic, `${page.ad} yakında`, d[page.id] || "Bu bölüm hazırlanıyor.");
  }
};


/* =====================================================================
   ARAÇLAR (F2)
   ===================================================================== */
const VEHICLE_TYPES = [{ id: "otomobil", ad: "Otomobil" }, { id: "motosiklet", ad: "Motosiklet" }, { id: "ticari", ad: "Ticari" }];
const FUELS = ["Benzin", "Dizel", "LPG", "Benzin + LPG", "Hibrit", "Elektrik"];
const NUM = n => new Intl.NumberFormat("tr-TR").format(n);
const trUp = s => String(s).toLocaleUpperCase("tr-TR");
const trLow = s => String(s ?? "").toLocaleLowerCase("tr-TR");
const typeAd = t => (VEHICLE_TYPES.find(x => x.id === t) || VEHICLE_TYPES[0]).ad;

function normalizePlate(raw) {
  const s = trUp(raw).replace(/[\s\-.]/g, "");
  const m = s.match(/^(\d{2})([A-Z]{1,3})(\d{2,5})$/);
  return m ? `${m[1]} ${m[2]} ${m[3]}` : null;
}
/* Uygulamadaki kurallarla aynı kontrol */
function validateVehicle(f, editId) {
  const e = {}; const yearNow = new Date().getFullYear();
  if (!f.marka) e.marka = "Marka boş bırakılamaz."; else if (f.marka.length > 30) e.marka = "Marka en fazla 30 karakter olabilir.";
  if (!f.model) e.model = "Model boş bırakılamaz."; else if (f.model.length > 40) e.model = "Model en fazla 40 karakter olabilir.";
  const yil = parseInt(f.yil, 10);
  if (!f.yil && f.yil !== 0) e.yil = "Model yılı girin."; else if (!(yil >= 1950 && yil <= yearNow + 1)) e.yil = `1950 ile ${yearNow + 1} arasında bir yıl girin.`;
  if (!f.plaka) e.plaka = "Plaka boş bırakılamaz.";
  else {
    const raw = trUp(f.plaka).replace(/[\s\-.]/g, "");
    if (/[ÇĞİÖŞÜ]/.test(raw)) e.plaka = "Plakada Türkçe karakter kullanılmaz.";
    else {
      const p = normalizePlate(f.plaka), il = p ? +p.slice(0, 2) : 0;
      if (!p) e.plaka = "Plaka biçimi hatalı. Örnek: 34 ABC 123";
      else if (il < 1 || il > 81) e.plaka = "İl kodu 01 ile 81 arasında olmalı.";
      else if ((S.vehicles || []).some(v => v.plaka === p && v.id !== editId)) e.plaka = "Bu plaka zaten kayıtlı.";
    }
  }
  const km = String(f.km ?? "").replace(/[.\s]/g, "");
  if (km === "") e.km = "Kilometre girin."; else if (!/^\d+$/.test(km)) e.km = "Kilometre yalnızca rakam olmalı."; else if (+km > 3000000) e.km = "Kilometre çok yüksek görünüyor.";
  if (!FUELS.includes(f.yakit)) e.yakit = "Yakıt türü seçin.";
  if (!VEHICLE_TYPES.some(t => t.id === f.tur)) e.tur = "Araç türü seçin.";
  const motor = String(f.motor ?? "").trim();
  if (motor && f.yakit !== "Elektrik" && !(/^\d+$/.test(motor) && +motor >= 49 && +motor <= 9000)) e.motor = "Motor hacmi 49–9000 cc olmalı.";
  return e;
}
const toVehicle = f => ({
  tur: f.tur, marka: String(f.marka).trim(), model: String(f.model).trim(), yil: parseInt(f.yil, 10),
  plaka: normalizePlate(f.plaka), km: +String(f.km).replace(/[.\s]/g, ""), yakit: f.yakit,
  motor: f.yakit === "Elektrik" || !String(f.motor ?? "").trim() ? null : +f.motor
});
const drivers = () => (S.members || []).filter(m => m.role === "driver" && m.active);
const driverName = uid => { const m = (S.members || []).find(x => x.uid === uid); return m ? (m.name || m.email) : ""; };

/* ---------- Pencere (modal) ---------- */
function openModal(html, wide) {
  closeModal();
  const w = document.createElement("div"); w.className = "modal-wrap"; w.id = "modal";
  w.innerHTML = `<div class="modal ${wide ? "wide" : ""}" role="dialog">${html}</div>`;
  w.addEventListener("mousedown", e => { if (e.target === w) closeModal(); });
  document.body.appendChild(w);
  w.querySelectorAll("[data-close]").forEach(b => b.addEventListener("click", closeModal));
  return w;
}
function closeModal() { const m = $("#modal"); if (m) m.remove(); }
document.addEventListener("keydown", e => { if (e.key === "Escape") closeModal(); });

Pages.vehicles = function () {
  $("#page").innerHTML = `
    <div class="toolbar">
      <div class="search">${I.search}<input class="input" id="vq" placeholder="Plaka, marka veya şoför ara"></div>
      <div class="tb-actions">
        <button class="btn btn-ghost" id="vExp">${I.down} Excel'e indir</button>
        <button class="btn btn-ghost" id="vImp">${I.up} Excel'den aktar</button>
        <button class="btn btn-primary" id="vAdd">${I.plus} Araç ekle</button>
      </div>
    </div>
    <div class="card" style="padding:0;overflow:hidden"><div id="vList"><div style="padding:30px"><div class="spinner"></div></div></div></div>`;
  $("#vAdd").addEventListener("click", () => vehicleModal());
  $("#vImp").addEventListener("click", importModal);
  $("#vExp").addEventListener("click", exportVehicles);
  $("#vq").addEventListener("input", renderVehicleList);
  Promise.all([api.listVehicles(S.company.id), api.listMembers(S.company.id)])
    .then(([vs, ms]) => { S.vehicles = vs; S.members = ms; renderVehicleList(); })
    .catch(err => { $("#vList").innerHTML = `<div class="alert err" style="margin:16px">${esc(errMsg(err))}</div>`; });
};

function renderVehicleList() {
  const box = $("#vList"); if (!box) return;
  const q = trLow($("#vq") ? $("#vq").value.trim() : "");
  const all = (S.vehicles || []).slice().sort((a, b) => a.plaka.localeCompare(b.plaka, "tr"));
  const list = q ? all.filter(v => trLow(`${v.plaka} ${v.plaka.replace(/\s/g, "")} ${v.marka} ${v.model} ${driverName(v.assignedUid)}`).includes(q)) : all;
  if (!all.length) {
    box.innerHTML = `<div class="empty"><div class="ic">${I.car}</div><h3>Henüz araç yok</h3>
      <p>Araçlarını tek tek ekleyebilir veya Excel'den toplu olarak aktarabilirsin.</p>
      <div style="display:flex;gap:10px;justify-content:center;margin-top:18px;flex-wrap:wrap">
        <button class="btn btn-ghost" data-act="imp">${I.up} Excel'den aktar</button><button class="btn btn-primary" data-act="add">${I.plus} Araç ekle</button></div></div>`;
    box.querySelector('[data-act="imp"]').addEventListener("click", importModal);
    box.querySelector('[data-act="add"]').addEventListener("click", () => vehicleModal());
    return;
  }
  const ds = drivers();
  box.innerHTML = `<div style="overflow-x:auto"><table class="tbl vt">
    <thead><tr><th>Plaka</th><th>Araç</th><th class="hm">Tür</th><th class="hm">Yakıt</th><th class="num">Km</th><th>Şoför</th><th></th></tr></thead>
    <tbody>${list.map(v => `<tr>
      <td><span class="plate"><i>TR</i><b>${esc(v.plaka)}</b></span></td>
      <td><b style="font-weight:600">${esc(v.marka)} ${esc(v.model)}</b><div class="muted" style="font-size:12.5px">${esc(v.yil)}</div></td>
      <td class="hm">${typeAd(v.tur)}</td><td class="hm">${esc(v.yakit)}</td>
      <td class="num">${NUM(v.km)}</td>
      <td><select class="input sel" data-assign="${v.id}">
        <option value="">— Atanmadı —</option>
        ${ds.map(d => `<option value="${d.uid}" ${v.assignedUid === d.uid ? "selected" : ""}>${esc(d.name || d.email)}</option>`).join("")}
        ${v.assignedUid && !ds.some(d => d.uid === v.assignedUid) ? `<option value="${v.assignedUid}" selected>(Ayrılmış şoför)</option>` : ""}
      </select></td>
      <td class="acts"><button class="icon" title="Düzenle" data-edit="${v.id}">${I.edit}</button><button class="icon danger" title="Sil" data-del="${v.id}">${I.trash}</button></td>
    </tr>`).join("")}</tbody></table></div>
    <div class="tfoot">${list.length === all.length ? `${all.length} araç` : `${list.length} / ${all.length} araç`}${ds.length ? "" : ` · <span class="muted">Şoför atamak için önce şoförlerin katılım koduyla katılmalı.</span>`}</div>`;
  box.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => vehicleModal(S.vehicles.find(v => v.id === b.dataset.edit))));
  box.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", () => deleteVehicleModal(S.vehicles.find(v => v.id === b.dataset.del))));
  box.querySelectorAll("[data-assign]").forEach(sel => sel.addEventListener("change", async () => {
    const v = S.vehicles.find(x => x.id === sel.dataset.assign); const newUid = sel.value || null, oldUid = v.assignedUid || null;
    sel.disabled = true;
    try {
      await api.assignVehicle(S.company.id, v.id, newUid, oldUid);
      v.assignedUid = newUid;
      (S.members || []).forEach(m => {
        m.vehicleIds = (m.vehicleIds || []).filter(id => id !== v.id);
        if (m.uid === newUid) m.vehicleIds.push(v.id);
      });
      toast(newUid ? `${v.plaka} → ${driverName(newUid)} atandı.` : `${v.plaka} için atama kaldırıldı.`);
    } catch (err) { sel.value = oldUid || ""; toast(errMsg(err), true); }
    finally { sel.disabled = false; }
  }));
}

function vehicleModal(v) {
  const d = v || { tur: "otomobil", marka: "", model: "", yil: "", plaka: "", km: "", yakit: "", motor: "" };
  const w = openModal(`
    <div class="m-head"><h3>${v ? "Aracı düzenle" : "Yeni araç"}</h3><button class="icon" data-close>${I.x}</button></div>
    <form id="vf" novalidate>
      <div class="g2">
        <div class="field" data-f="tur"><label>Araç türü *</label><select class="input" name="tur">${VEHICLE_TYPES.map(t => `<option value="${t.id}" ${d.tur === t.id ? "selected" : ""}>${t.ad}</option>`).join("")}</select><div class="msg"></div></div>
        <div class="field" data-f="plaka"><label>Plaka *</label><input class="input" name="plaka" maxlength="12" placeholder="34 ABC 123" style="text-transform:uppercase" value="${esc(d.plaka)}"><div class="msg"></div></div>
        <div class="field" data-f="marka"><label>Marka *</label><input class="input" name="marka" maxlength="30" value="${esc(d.marka)}"><div class="msg"></div></div>
        <div class="field" data-f="model"><label>Model *</label><input class="input" name="model" maxlength="40" value="${esc(d.model)}"><div class="msg"></div></div>
        <div class="field" data-f="yil"><label>Model yılı *</label><input class="input" name="yil" inputmode="numeric" maxlength="4" value="${esc(d.yil)}"><div class="msg"></div></div>
        <div class="field" data-f="km"><label>Güncel kilometre *</label><input class="input" name="km" inputmode="numeric" value="${d.km !== "" ? NUM(d.km) : ""}"><div class="msg"></div></div>
        <div class="field" data-f="yakit"><label>Yakıt türü *</label><select class="input" name="yakit"><option value="">Seçin</option>${FUELS.map(f => `<option ${d.yakit === f ? "selected" : ""}>${f}</option>`).join("")}</select><div class="msg"></div></div>
        <div class="field" data-f="motor"><label>Motor hacmi (cc)</label><input class="input" name="motor" inputmode="numeric" maxlength="4" value="${d.motor ? esc(d.motor) : ""}"><div class="msg"></div></div>
      </div>
      <div class="m-foot"><button type="button" class="btn btn-ghost" data-close>Vazgeç</button><button type="submit" class="btn btn-primary">${v ? "Kaydet" : "Aracı ekle"}</button></div>
    </form>`);
  const f = w.querySelector("#vf");
  f.km.addEventListener("input", () => { const x = f.km.value.replace(/\D/g, "").slice(0, 7); f.km.value = x ? NUM(+x) : ""; });
  f.plaka.addEventListener("blur", () => { const p = normalizePlate(f.plaka.value); if (p) f.plaka.value = p; });
  f.addEventListener("input", e => { const el = f.querySelector(`[data-f="${e.target.name}"]`); if (el) el.classList.remove("err"); });
  setTimeout(() => f.plaka.focus(), 50);
  f.addEventListener("submit", e => {
    e.preventDefault();
    const val = { tur: f.tur.value, marka: f.marka.value.trim(), model: f.model.value.trim(), yil: f.yil.value.trim(), plaka: f.plaka.value.trim(), km: f.km.value.trim(), yakit: f.yakit.value, motor: f.motor.value.trim() };
    const errs = validateVehicle(val, v ? v.id : null);
    f.querySelectorAll(".field").forEach(el => { const k = el.dataset.f; el.classList.toggle("err", !!errs[k]); el.querySelector(".msg").textContent = errs[k] || ""; });
    if (Object.keys(errs).length) return;
    busy(f.querySelector('button[type=submit]'), async () => {
      try {
        const data = toVehicle(val);
        if (v) { await api.updateVehicle(S.company.id, v.id, data); Object.assign(v, data); toast("Araç güncellendi."); }
        else { const id = await api.addVehicle(S.company.id, data); S.vehicles.push({ id, ...data, assignedUid: null }); toast("Araç eklendi."); }
        closeModal(); renderVehicleList();
      } catch (err) { toast(errMsg(err), true); }
    });
  });
}

function deleteVehicleModal(v) {
  const w = openModal(`
    <div class="m-head"><h3>Araç silinsin mi?</h3><button class="icon" data-close>${I.x}</button></div>
    <p class="muted"><b style="color:var(--text)">${esc(v.plaka)} – ${esc(v.marka)} ${esc(v.model)}</b> silinecek.${v.assignedUid ? ` Şoför ataması (${esc(driverName(v.assignedUid))}) da kaldırılacak.` : ""} Bu işlem geri alınamaz.</p>
    <div class="m-foot"><button class="btn btn-ghost" data-close>Vazgeç</button><button class="btn btn-danger" id="delOk">${I.trash} Sil</button></div>`);
  w.querySelector("#delOk").addEventListener("click", e => busy(e.currentTarget, async () => {
    try {
      await api.deleteVehicle(S.company.id, v.id, v.assignedUid || null);
      S.vehicles = S.vehicles.filter(x => x.id !== v.id);
      (S.members || []).forEach(m => { m.vehicleIds = (m.vehicleIds || []).filter(id => id !== v.id); });
      closeModal(); renderVehicleList(); toast("Araç silindi.");
    } catch (err) { toast(errMsg(err), true); }
  }));
}

/* ---------- Excel ---------- */
const XL_HEAD = ["Plaka", "Marka", "Model", "Model Yılı", "Kilometre", "Yakıt Türü", "Araç Türü", "Motor Hacmi (cc)"];
function templateDownload() {
  const blob = writeXlsx([
    { name: "Araçlar", widths: [14, 16, 18, 12, 13, 15, 13, 17], rows: [XL_HEAD, ["34 ABC 123", "Renault", "Clio", 2019, 85000, "Benzin", "Otomobil", 999], ["06 BSY 41", "Ford", "Transit Custom", 2021, 64200, "Dizel", "Ticari", ""]] },
    { name: "Açıklama", widths: [20, 70], rows: [["Sütun", "Nasıl doldurulur?"],
      ["Plaka *", "Örn. 34 ABC 123 (boşluklar önemli değil)"], ["Marka *", "Örn. Renault"], ["Model *", "Örn. Clio"],
      ["Model Yılı *", "Örn. 2019"], ["Kilometre *", "Sadece rakam. Örn. 85000"],
      ["Yakıt Türü *", FUELS.join(", ")], ["Araç Türü", "Otomobil, Motosiklet veya Ticari (boşsa Otomobil)"],
      ["Motor Hacmi (cc)", "İsteğe bağlı. Örn. 1461"], ["Not", "Örnek satırları silip kendi araçlarını yaz. * işaretli sütunlar zorunlu."]] }
  ]);
  download(blob, "driventa-arac-sablonu.xlsx");
}
function mapFuel(x) {
  const k = trLow(x).replace(/[\s+\-_/]/g, "");
  if (!k) return "";
  if (k === "benzinlpg" || k === "lpgbenzin") return "Benzin + LPG";
  if (k.startsWith("benzin")) return "Benzin";
  if (["dizel", "diesel", "motorin", "mazot"].includes(k)) return "Dizel";
  if (k === "lpg") return "LPG";
  if (k.startsWith("hibrit") || k.startsWith("hybrid")) return "Hibrit";
  if (k.startsWith("elektrik") || k === "ev") return "Elektrik";
  return String(x).trim();
}
function mapType(x) {
  const k = trLow(x).trim();
  if (!k) return "otomobil";
  if (["otomobil", "binek", "araba", "oto"].includes(k)) return "otomobil";
  if (["motosiklet", "motor", "moto", "motorsiklet"].includes(k)) return "motosiklet";
  if (["ticari", "kamyonet", "panelvan", "minibüs", "minibus", "kamyon", "hafif ticari", "pikap"].includes(k)) return "ticari";
  return "?";
}
const HEAD_KEYS = { plaka: "plaka", marka: "marka", model: "model", modelyılı: "yil", modelyili: "yil", yıl: "yil", yil: "yil",
  kilometre: "km", km: "km", güncelkilometre: "km", yakıttürü: "yakit", yakitturu: "yakit", yakıt: "yakit", yakit: "yakit",
  araçtürü: "tur", aracturu: "tur", tür: "tur", tur: "tur", tip: "tur", motorhacmi: "motor", motorhacmicc: "motor", motor: "motor", cc: "motor" };

function parseRows(rows) {
  const headIdx = rows.findIndex(r => r && r.some(c => HEAD_KEYS[trLow(c).replace(/[^a-zçğıöşü]/g, "")] === "plaka"));
  if (headIdx < 0) return { error: "Dosyada \"Plaka\" sütunu bulunamadı. Lütfen şablonu kullan." };
  const map = rows[headIdx].map(c => HEAD_KEYS[trLow(c).replace(/[^a-zçğıöşü]/g, "")] || null);
  const need = ["plaka", "marka", "model", "yil", "km", "yakit"];
  const missing = need.filter(k => !map.includes(k));
  if (missing.length) return { error: `Eksik sütun: ${missing.map(k => ({ plaka: "Plaka", marka: "Marka", model: "Model", yil: "Model Yılı", km: "Kilometre", yakit: "Yakıt Türü" })[k]).join(", ")}` };
  const out = []; const seen = new Set();
  rows.slice(headIdx + 1).forEach((r, i) => {
    if (!r || r.every(c => String(c ?? "").trim() === "")) return;
    const o = {}; map.forEach((k, ci) => { if (k) o[k] = r[ci] ?? ""; });
    const f = { plaka: String(o.plaka ?? "").trim(), marka: String(o.marka ?? "").trim(), model: String(o.model ?? "").trim(),
      yil: String(o.yil ?? "").trim(), km: String(o.km ?? "").trim(), yakit: mapFuel(o.yakit), tur: mapType(o.tur ?? ""), motor: String(o.motor ?? "").trim() };
    const errs = validateVehicle(f, null);
    if (f.tur === "?") errs.tur = "Araç türü Otomobil, Motosiklet veya Ticari olmalı.";
    if (errs.yakit && f.yakit) errs.yakit = `"${f.yakit}" tanınmadı. ${FUELS.join(", ")} olmalı.`;
    const p = normalizePlate(f.plaka);
    if (p && seen.has(p) && !errs.plaka) errs.plaka = "Bu plaka dosyada birden fazla kez geçiyor.";
    if (p) seen.add(p);
    out.push({ line: headIdx + i + 2, f, errs, ok: !Object.keys(errs).length });
  });
  return { rows: out };
}

function importModal() {
  const w = openModal(`
    <div class="m-head"><h3>Excel'den araç aktar</h3><button class="icon" data-close>${I.x}</button></div>
    <ol class="steps" style="margin:0 0 16px">
      <li><span><a class="link" href="#" id="tpl">Şablonu indir</a> ve araçlarını doldur.</span></li>
      <li><span>Doldurduğun dosyayı (.xlsx veya .csv) aşağıdan seç.</span></li>
      <li><span>Önizlemeyi kontrol et ve aktar.</span></li>
    </ol>
    <label class="drop" id="drop"><input type="file" id="xf" accept=".xlsx,.csv" hidden>${I.up}<b>Dosya seç</b><span class="muted">veya buraya sürükle bırak</span></label>
    <div id="impRes"></div>`, true);
  w.querySelector("#tpl").addEventListener("click", e => { e.preventDefault(); templateDownload(); });
  const input = w.querySelector("#xf"), drop = w.querySelector("#drop");
  input.addEventListener("change", () => input.files[0] && handleFile(input.files[0]));
  drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("over"));
  drop.addEventListener("drop", e => { e.preventDefault(); drop.classList.remove("over"); const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
}

async function handleFile(file) {
  const res = $("#impRes");
  if (file.size > 5 * 1024 * 1024) { res.innerHTML = '<div class="alert err">Dosya çok büyük (en fazla 5 MB).</div>'; return; }
  res.innerHTML = '<div style="padding:20px"><div class="spinner"></div></div>';
  let rows;
  try {
    if (/\.csv$/i.test(file.name)) rows = readCsv(await file.text());
    else if (/\.xlsx$/i.test(file.name)) rows = await readXlsx(await file.arrayBuffer());
    else { res.innerHTML = '<div class="alert err">Lütfen .xlsx veya .csv dosyası seç. (Eski .xls dosyasını Excel\'de "Farklı kaydet → .xlsx" ile dönüştürebilirsin.)</div>'; return; }
  } catch (e) { console.error(e); res.innerHTML = '<div class="alert err">Dosya okunamadı. Dosyanın bozuk olmadığından ve .xlsx biçiminde kaydedildiğinden emin ol.</div>'; return; }
  const parsed = parseRows(rows);
  if (parsed.error) { res.innerHTML = `<div class="alert err">${esc(parsed.error)}</div>`; return; }
  const ok = parsed.rows.filter(r => r.ok), bad = parsed.rows.filter(r => !r.ok);
  if (!parsed.rows.length) { res.innerHTML = '<div class="alert err">Dosyada araç satırı bulunamadı.</div>'; return; }
  const LBL = { plaka: "Plaka", marka: "Marka", model: "Model", yil: "Yıl", km: "Km", yakit: "Yakıt", tur: "Tür", motor: "Motor" };
  res.innerHTML = `
    <div class="sum"><span class="pill ok">${ok.length} araç aktarılacak</span>${bad.length ? `<span class="pill bad">${bad.length} satırda hata</span>` : ""}</div>
    <div class="prev"><table class="tbl"><thead><tr><th>Satır</th><th>Plaka</th><th>Araç</th><th>Km</th><th>Yakıt</th><th>Durum</th></tr></thead><tbody>
      ${parsed.rows.map(r => `<tr class="${r.ok ? "" : "bad"}"><td class="muted">${r.line}</td><td style="white-space:nowrap">${esc(normalizePlate(r.f.plaka) || r.f.plaka || "–")}</td>
        <td>${esc(r.f.marka)} ${esc(r.f.model)} <span class="muted">${esc(r.f.yil)}</span></td><td>${esc(r.f.km)}</td><td>${esc(r.f.yakit)}</td>
        <td>${r.ok ? `<span class="okc">${I.ok}</span>` : `<span class="errc">${Object.entries(r.errs).map(([k, m]) => `<b>${LBL[k] || k}:</b> ${esc(m)}`).join("<br>")}</span>`}</td></tr>`).join("")}
    </tbody></table></div>
    ${bad.length ? '<p class="muted" style="font-size:13px;margin-top:10px">Hatalı satırlar atlanır. İstersen dosyayı düzeltip tekrar yükleyebilirsin.</p>' : ""}
    <div class="m-foot"><button class="btn btn-ghost" data-close2>Vazgeç</button><button class="btn btn-primary" id="doImp" ${ok.length ? "" : "disabled"}>${I.up} ${ok.length} aracı aktar</button></div>`;
  res.querySelector("[data-close2]").addEventListener("click", closeModal);
  const btn = res.querySelector("#doImp");
  btn.addEventListener("click", () => busy(btn, async () => {
    try {
      const list = ok.map(r => toVehicle(r.f));
      await api.bulkAddVehicles(S.company.id, list);
      S.vehicles = await api.listVehicles(S.company.id);
      closeModal(); renderVehicleList(); toast(`${list.length} araç aktarıldı.`);
    } catch (err) { toast(errMsg(err), true); }
  }));
}

function exportVehicles() {
  const vs = (S.vehicles || []).slice().sort((a, b) => a.plaka.localeCompare(b.plaka, "tr"));
  if (!vs.length) return toast("İndirilecek araç yok.", true);
  const rows = [[...XL_HEAD, "Atanan Şoför"], ...vs.map(v => [v.plaka, v.marka, v.model, v.yil, v.km, v.yakit, typeAd(v.tur), v.motor ?? "", driverName(v.assignedUid) || ""])];
  const d = new Date(); const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  download(writeXlsx([{ name: "Araçlar", widths: [14, 16, 18, 12, 13, 15, 13, 17, 22], rows }]), `driventa-araclar-${stamp}.xlsx`);
  toast("Excel dosyası indirildi.");
}

/* İki aşamalı onay: ilk basışta uyarır, ikinci basışta işlemi yapar */
function confirmInline(btn, label) {
  if (btn.dataset.armed) { delete btn.dataset.armed; btn.innerHTML = btn.dataset.orig; return true; }
  btn.dataset.armed = "1"; btn.dataset.orig = btn.innerHTML; btn.innerHTML = label;
  setTimeout(() => { if (btn.dataset.armed) { delete btn.dataset.armed; btn.innerHTML = btn.dataset.orig; } }, 3500);
  return false;
}

/* ---------- Yükleme akışı ---------- */
async function loadCompany(cid) {
  const [company, member] = await Promise.all([api.getCompany(cid), api.getMember(cid, S.user.uid)]);
  if (!company || !member || !member.active) { renderOnboarding(); return; }
  S.company = company; S.member = member;
  if (member.role === "driver") return renderDriverInfo();
  renderShell();
}

api.onUser(async user => {
  S.user = user; S.company = null; S.member = null; S.page = "overview";
  if (!user) return renderAuth("login");
  root.innerHTML = '<div class="center"><div class="spinner"></div></div>';
  try {
    S.profile = await api.getProfile(user.uid);
    if (S.profile && S.profile.companyId) await loadCompany(S.profile.companyId);
    else renderOnboarding();
  } catch (err) {
    root.innerHTML = `<div class="center"><div class="card"><div class="alert err">${esc(errMsg(err))}</div>
      <button class="btn btn-ghost" id="retry">Tekrar dene</button> <button class="btn btn-ghost" id="lo">Çıkış yap</button></div></div>`;
    $("#retry").addEventListener("click", () => location.reload());
    $("#lo").addEventListener("click", () => api.signOut());
  }
});
