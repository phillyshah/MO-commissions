// Passkey ceremonies for the approvals pages. No framework. Every button that uses this file has a
// non-JS fallback or a plain explanation when the browser cannot do passkeys.
(function () {
  "use strict";

  function b64urlToBuf(s) {
    s = s.replace(/-/g, "+").replace(/_/g, "/");
    while (s.length % 4) s += "=";
    const bin = atob(s); const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    return buf.buffer;
  }
  function bufToB64url(buf) {
    const bytes = new Uint8Array(buf); let bin = "";
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function status(text) {
    const el = document.getElementById("webauthn-status");
    if (!el) { if (text) window.alert(text); return; }
    el.textContent = text || ""; el.hidden = !text;
    if (text) el.scrollIntoView({ block: "nearest" });
  }

  async function postJSON(url, body) {
    const r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, credentials: "same-origin",
                                 body: JSON.stringify(body || {}) });
    let data = {};
    try { data = await r.json(); } catch (e) { /* not JSON */ }
    if (r.status === 401 || r.status === 404) throw new Error("Your sign-in has ended. Please sign in again; nothing was lost.");
    if (!r.ok || data.ok === false) throw new Error(data.error || "Something went wrong. Nothing was lost — try again.");
    return data;
  }

  function friendly(err) {
    const name = err && err.name;
    if (name === "NotAllowedError") return "The passkey prompt was cancelled or timed out. Nothing was signed — try again when ready.";
    if (name === "InvalidStateError") return "This device already has a passkey for you. Remove it under My passkeys before adding another.";
    if (name === "SecurityError") return "This page is not being served from the portal's own address, so the passkey cannot be used.";
    if (name === "NotSupportedError") return "This browser cannot use passkeys. Try your phone, or a current version of Chrome, Edge or Safari.";
    return (err && err.message) || "Something went wrong. Nothing was lost — try again.";
  }

  // Decode the server's options (base64url fields) into the ArrayBuffers the browser API wants.
  function creationOptions(o) {
    o.challenge = b64urlToBuf(o.challenge);
    o.user.id = b64urlToBuf(o.user.id);
    (o.excludeCredentials || []).forEach((c) => { c.id = b64urlToBuf(c.id); });
    return o;
  }
  function requestOptions(o) {
    o.challenge = b64urlToBuf(o.challenge);
    (o.allowCredentials || []).forEach((c) => { c.id = b64urlToBuf(c.id); });
    return o;
  }
  function serialize(cred) {
    const r = cred.response;
    const out = { id: cred.id, rawId: bufToB64url(cred.rawId), type: cred.type,
                  clientExtensionResults: cred.getClientExtensionResults ? cred.getClientExtensionResults() : {},
                  response: { clientDataJSON: bufToB64url(r.clientDataJSON) } };
    if (r.attestationObject) {
      out.response.attestationObject = bufToB64url(r.attestationObject);
      if (r.getTransports) out.response.transports = r.getTransports();
    } else {
      out.response.authenticatorData = bufToB64url(r.authenticatorData);
      out.response.signature = bufToB64url(r.signature);
      if (r.userHandle) out.response.userHandle = bufToB64url(r.userHandle);
    }
    if (cred.authenticatorAttachment) out.authenticatorAttachment = cred.authenticatorAttachment;
    return out;
  }

  async function register(btn) {
    const code = (document.getElementById(btn.dataset.codeFrom) || {}).value || "";
    const label = (document.getElementById(btn.dataset.labelFrom) || {}).value || "";
    if (!code.trim()) throw new Error("Enter the enrollment code the Controller gave you.");
    const { options } = await postJSON("/passkeys/webauthn/register/options", { code: code });
    const cred = await navigator.credentials.create({ publicKey: creationOptions(options) });
    const done = await postJSON("/passkeys/webauthn/register/verify", { code: code, label: label, credential: serialize(cred) });
    window.location.assign(done.redirect || "/passkeys");
  }

  async function session(btn) {
    const { options } = await postJSON("/approvals/webauthn/session/options", {});
    const cred = await navigator.credentials.get({ publicKey: requestOptions(options) });
    const done = await postJSON("/approvals/webauthn/session/verify", { credential: serialize(cred), next: btn.dataset.next || "/approvals" });
    window.location.assign(done.redirect || "/approvals");
  }

  async function sign(btn) {
    const requestId = btn.dataset.request, decision = btn.dataset.decision || "approve";
    let reason = "";
    if (btn.dataset.reasonFrom) {
      const form = document.getElementById(btn.dataset.reasonFrom);
      const field = form && form.querySelector("[name=reason]");
      reason = field ? field.value.trim() : "";
      if (decision === "reject" && reason.length < 3) throw new Error("Give a reason for rejecting (a few words is enough).");
    }
    const { options } = await postJSON("/approvals/webauthn/sign/" + requestId + "/options", { decision: decision });
    const cred = await navigator.credentials.get({ publicKey: requestOptions(options) });
    const done = await postJSON("/approvals/webauthn/sign/" + requestId + "/verify", { credential: serialize(cred), decision: decision, reason: reason });
    window.location.assign(done.redirect || "/approvals");
  }

  document.addEventListener("click", async (e) => {
    const open = e.target.closest("[data-open-dialog]");
    if (open) { const d = document.getElementById(open.dataset.openDialog); if (d) { d.returnTrigger = open; d.showModal(); } return; }
    const btn = e.target.closest("[data-webauthn]");
    if (!btn) return;
    e.preventDefault();
    if (!window.PublicKeyCredential || !navigator.credentials) {
      status("This browser cannot use passkeys. Try your phone, or a current version of Chrome, Edge or Safari."); return;
    }
    const label = btn.textContent; btn.disabled = true; btn.textContent = "Waiting for your passkey…"; status("");
    try {
      if (btn.dataset.webauthn === "register") await register(btn);
      else if (btn.dataset.webauthn === "session") await session(btn);
      else if (btn.dataset.webauthn === "sign") await sign(btn);
    } catch (err) {
      status(friendly(err));
    } finally {
      btn.disabled = false; btn.textContent = label;
    }
  });
})();
