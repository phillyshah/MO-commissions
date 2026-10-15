// Small progressive enhancements. No framework; every control works without this file.
document.addEventListener("click", (e) => {
  const t = e.target.closest("[data-toggle-password]");
  if (t) {
    const input = document.getElementById(t.dataset.togglePassword);
    const show = input.type === "password";
    input.type = show ? "text" : "password";
    t.textContent = show ? "Hide" : "Show";
    t.setAttribute("aria-pressed", String(show));
  }
  const h = e.target.closest("[data-help]");
  if (h) {
    const dlg = document.getElementById(h.dataset.help);
    if (dlg) { dlg.returnTrigger = h; dlg.showModal(); }
  }
  const c = e.target.closest("[data-close]");
  if (c) { const dlg = c.closest("dialog"); if (dlg) { dlg.close(); if (dlg.returnTrigger) dlg.returnTrigger.focus(); } }
  const row = e.target.closest("tr.clickable[data-href]");
  if (row && !e.target.closest("a,button,input")) window.location.assign(row.dataset.href);
});
document.querySelectorAll("form[data-confirm]").forEach((f) => {
  f.addEventListener("submit", (e) => { if (!window.confirm(f.dataset.confirm)) e.preventDefault(); });
});
// Buttons never stick on "Working…": the busy flag is cleared if the page is still here after 15s.
document.querySelectorAll("form").forEach((f) => {
  f.addEventListener("submit", () => {
    const b = f.querySelector("button[type=submit], button:not([type])");
    if (!b) return;
    b.dataset.label = b.textContent; b.textContent = "Working…"; b.disabled = true;
    setTimeout(() => { b.textContent = b.dataset.label; b.disabled = false; }, 15000);
  });
});
