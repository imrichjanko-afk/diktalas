/* Diktálás külön ablakban.
 * Az Outlook a beágyazott panelben nem mindig engedi a mikrofont; ez az oldal
 * az Office párbeszédablakában (önálló ablakban) fut, ahol a mikrofon a
 * szokásos módon engedélyezhető. A szöveget messageParent-tel adja vissza.
 */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const LANG_NAMES = { "hu-HU": "Magyar", "sk-SK": "Slovenčina" };
  const params = new URLSearchParams(location.search);
  const lang = LANG_NAMES[params.get("lang")] ? params.get("lang") : "hu-HU";
  let creds = {
    key: localStorage.getItem("diktalas.speechKey") || "",
    region: localStorage.getItem("diktalas.speechRegion") || "",
  };
  let recognizer = null;

  function setStatus(msg, kind) {
    $("status").textContent = msg || "";
    $("status").className = "status" + (kind ? " " + kind : "");
  }

  function inDialog() {
    return typeof Office !== "undefined" && Office.context && Office.context.ui &&
      typeof Office.context.ui.messageParent === "function";
  }

  function send(obj) {
    if (inDialog()) Office.context.ui.messageParent(JSON.stringify(obj));
  }

  function setRecordingUI(on) {
    $("micBtn").classList.toggle("recording", on);
    $("micBtn").setAttribute("aria-pressed", String(on));
    $("micLabel").textContent = on ? "Diktálás leállítása" : "Diktálás indítása";
    if (!on) $("interim").textContent = "";
  }

  function appendText(chunk) {
    chunk = chunk.trim();
    if (!chunk) return;
    const t = $("text");
    const sep = t.value && !/\s$/.test(t.value) ? " " : "";
    t.value += sep + chunk;
    t.scrollTop = t.scrollHeight;
  }

  async function start() {
    if (!creds.key || !creds.region) {
      setStatus("Hiányzik a kulcs vagy a régió – mentse el a panel Beállításaiban, majd nyissa meg újra ezt az ablakot.", "error");
      return;
    }
    if (typeof SpeechSDK === "undefined") {
      setStatus("A beszédfelismerő könyvtár nem töltődött be.", "error");
      return;
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      stream.getTracks().forEach((t) => t.stop());
    } catch (e) {
      setStatus("Mikrofon nem érhető el ebben az ablakban sem: " + (e.message || e) +
        " [" + (e.name || "") + "]", "error");
      return;
    }
    const cfg = SpeechSDK.SpeechConfig.fromSubscription(creds.key, creds.region);
    cfg.speechRecognitionLanguage = lang;
    cfg.setProperty(SpeechSDK.PropertyId.Speech_SegmentationSilenceTimeoutMs, "1200");
    recognizer = new SpeechSDK.SpeechRecognizer(cfg, SpeechSDK.AudioConfig.fromDefaultMicrophoneInput());
    recognizer.recognizing = (_s, e) => { $("interim").textContent = e.result.text; };
    recognizer.recognized = (_s, e) => {
      if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
        appendText(e.result.text);
        $("interim").textContent = "";
      }
    };
    recognizer.canceled = (_s, e) => {
      if (e.reason === SpeechSDK.CancellationReason.Error) setStatus("Felismerési hiba: " + e.errorDetails, "error");
      stop();
    };
    recognizer.startContinuousRecognitionAsync(
      () => { setRecordingUI(true); setStatus("Hallgatom… beszéljen nyugodtan.", ""); },
      (err) => { setStatus("Nem sikerült elindítani: " + err, "error"); stop(); }
    );
  }

  function stop(after) {
    const r = recognizer;
    recognizer = null;
    setRecordingUI(false);
    if (!r) { if (after) after(); return; }
    r.stopContinuousRecognitionAsync(
      () => { r.close(); if (after) after(); },
      () => { r.close(); if (after) after(); }
    );
  }

  function done() {
    stop(() => {
      const text = $("text").value.trim();
      if (!inDialog()) { setStatus("Ez az ablak csak az Outlookból megnyitva tud szöveget visszaadni.", "error"); return; }
      send({ type: "text", text: text });
    });
  }

  function init() {
    $("langInfo").textContent = "Nyelv: " + LANG_NAMES[lang];
    $("micBtn").addEventListener("click", () => (recognizer ? stop() : start()));
    $("doneBtn").addEventListener("click", done);
    // Ha a párbeszédablak nem látja a panel tárolóját, a panel küldi át a kulcsot.
    if (inDialog() && Office.context.ui.addHandlerAsync && Office.EventType &&
        Office.EventType.DialogParentMessageReceived) {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        try {
          const msg = JSON.parse(arg.message);
          if (msg.type === "creds" && msg.key && msg.region) {
            creds = { key: msg.key, region: msg.region };
            setStatus("Kész. Nyomja meg a „Diktálás indítása” gombot.", "ok");
          }
        } catch (_) { /* figyelmen kívül hagyva */ }
      });
    }
    send({ type: "ready" });
    if (creds.key && creds.region) setStatus("Kész. Nyomja meg a „Diktálás indítása” gombot.", "ok");
  }

  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(init);
  else document.addEventListener("DOMContentLoaded", init);
})();
