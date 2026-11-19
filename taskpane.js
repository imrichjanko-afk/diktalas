/* Diktálás – Outlook bővítmény
 * Magyar / szlovák diktálás (Azure Speech), német / angol fordítás (Azure Translator),
 * jóváhagyás után beszúrás az épp írt e-mailbe.
 */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const els = {
    lang: $("lang"), micBtn: $("micBtn"), micLabel: $("micLabel"), interim: $("interim"),
    text: $("text"), toDe: $("toDe"), toEn: $("toEn"), undo: $("undo"), clear: $("clear"),
    insertBtn: $("insertBtn"), status: $("status"), settings: $("settings"),
    speechKey: $("speechKey"), speechRegion: $("speechRegion"),
    trKey: $("trKey"), trRegion: $("trRegion"), saveSettings: $("saveSettings"),
  };

  const SETTING_KEYS = ["speechKey", "speechRegion", "trKey", "trRegion", "lang"];
  let recognizer = null;
  let history = [];

  // ---------- Segédfüggvények ----------
  function setStatus(msg, kind) {
    els.status.textContent = msg || "";
    els.status.className = "status" + (kind ? " " + kind : "");
  }

  function inOutlook() {
    return typeof Office !== "undefined" && Office.context && Office.context.mailbox;
  }

  function pushHistory() {
    history.push(els.text.value);
    if (history.length > 30) history.shift();
    els.undo.disabled = false;
  }

  // ---------- Beállítások (roamingSettings az Outlookban, különben localStorage) ----------
  function readSetting(name) {
    if (inOutlook() && Office.context.roamingSettings) {
      const v = Office.context.roamingSettings.get(name);
      if (v !== undefined && v !== null) return v;
    }
    return localStorage.getItem("diktalas." + name) || "";
  }

  function loadSettings() {
    SETTING_KEYS.forEach((k) => {
      const v = readSetting(k);
      if (v && els[k]) els[k].value = v;
    });
    if (!els.speechKey.value || !els.speechRegion.value) {
      els.settings.open = true;
      setStatus("Először adja meg az Azure-kulcsokat a Beállításokban.", "error");
    }
  }

  function saveSettings() {
    const values = {};
    SETTING_KEYS.forEach((k) => { values[k] = (els[k].value || "").trim(); });
    Object.entries(values).forEach(([k, v]) => localStorage.setItem("diktalas." + k, v));
    if (inOutlook() && Office.context.roamingSettings) {
      const rs = Office.context.roamingSettings;
      Object.entries(values).forEach(([k, v]) => rs.set(k, v));
      rs.saveAsync((r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) setStatus("Beállítások elmentve.", "ok");
        else setStatus("Mentési hiba: " + r.error.message, "error");
      });
    } else {
      setStatus("Beállítások elmentve (helyileg).", "ok");
    }
  }

  // ---------- Mikrofon-engedély (új Outlook / Outlook a weben) ----------
  // Az Outlook a bővítmény keretében külön kéri az engedélyt. Ha a felhasználó
  // korábban elutasította, újra fel kell kínálni a párbeszédablakot.
  async function requestOfficeMicPermission() {
    const api = typeof Office !== "undefined" && Office.devicePermission;
    const types = typeof Office !== "undefined" && Office.DevicePermissionType;
    if (!api || !types) return "nincs-api";
    if (typeof api.requestPermissionsAsync === "function") {
      // Az Outlook ezt a függvényt visszahívással (callback) várja, nem Promise-szal.
      const granted = await new Promise((resolve, reject) => {
        api.requestPermissionsAsync([types.microphone], (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(result.error || new Error("ismeretlen hiba"));
        });
      });
      return granted ? "megadva" : "elutasitva";
    }
    if (typeof api.requestPermissions === "function") {
      // Ennél a változatnál az Outlook engedélyezés után újratölti a panelt.
      await api.requestPermissions([types.microphone]);
      return "megadva";
    }
    return "nincs-api";
  }

  async function ensureMicPermission() {
    let officeResult = "nincs-api";
    try {
      officeResult = await requestOfficeMicPermission();
    } catch (e) {
      officeResult = "hiba: " + ((e && (e.message || e.code)) || e);
    }
    if (officeResult === "elutasitva") {
      throw new Error("Az Outlook engedélykérő ablakában a mikrofon nem lett engedélyezve. " +
        "Nyomja meg újra a Diktálás indítása gombot, és válassza az „Engedélyezés” lehetőséget.");
    }
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      throw new Error("Ez a felület nem támogatja a mikrofont.");
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      stream.getTracks().forEach((t) => t.stop());
    } catch (e) {
      const name = (e && e.name) || "";
      let hint = "";
      if (name === "NotAllowedError") {
        hint = officeResult === "megadva"
          ? " Az engedélyt megadta: zárja be a Diktálás panelt, nyissa meg újra, és próbálja ismét."
          : " Az Outlook nem adott engedélyt a bővítménynek (" + officeResult + ").";
      } else if (name === "NotFoundError") {
        hint = " Nem található mikrofon – ellenőrizze, hogy csatlakoztatva van-e.";
      } else if (name === "NotReadableError") {
        hint = " A mikrofont egy másik program használja (pl. Teams-hívás).";
      }
      throw new Error((e && e.message) + " [" + name + "]" + hint);
    }
  }

  // ---------- Diktálás ----------
  function appendText(chunk) {
    chunk = chunk.trim();
    if (!chunk) return;
    const cur = els.text.value;
    const sep = cur && !/\s$/.test(cur) ? " " : "";
    els.text.value = cur + sep + chunk;
    els.text.scrollTop = els.text.scrollHeight;
  }

  function setRecordingUI(on) {
    els.micBtn.classList.toggle("recording", on);
    els.micBtn.setAttribute("aria-pressed", String(on));
    els.micLabel.textContent = on ? "Diktálás leállítása" : "Diktálás indítása";
    els.lang.disabled = on;
    if (!on) els.interim.textContent = "";
  }

  async function startDictation() {
    const key = els.speechKey.value.trim();
    const region = els.speechRegion.value.trim();
    if (!key || !region) {
      els.settings.open = true;
      setStatus("Hiányzik az Azure Speech kulcs vagy régió.", "error");
      return;
    }
    if (typeof SpeechSDK === "undefined") {
      setStatus("A beszédfelismerő könyvtár nem töltődött be (internetkapcsolat?).", "error");
      return;
    }
    try {
      await ensureMicPermission();
    } catch (e) {
      setStatus("Mikrofon nem érhető el: " + (e.message || e), "error");
      return;
    }

    const speechConfig = SpeechSDK.SpeechConfig.fromSubscription(key, region);
    speechConfig.speechRecognitionLanguage = els.lang.value;
    speechConfig.setProperty(SpeechSDK.PropertyId.Speech_SegmentationSilenceTimeoutMs, "1200");
    const audioConfig = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();
    recognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);

    let historySaved = false;
    recognizer.recognizing = (_s, e) => { els.interim.textContent = e.result.text; };
    recognizer.recognized = (_s, e) => {
      if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
        if (!historySaved) { pushHistory(); historySaved = true; }
        appendText(e.result.text);
        els.interim.textContent = "";
      }
    };
    recognizer.canceled = (_s, e) => {
      if (e.reason === SpeechSDK.CancellationReason.Error) {
        setStatus("Felismerési hiba: " + e.errorDetails, "error");
      }
      stopDictation();
    };

    recognizer.startContinuousRecognitionAsync(
      () => { setRecordingUI(true); setStatus("Hallgatom… beszéljen nyugodtan.", ""); },
      (err) => { setStatus("Nem sikerült elindítani: " + err, "error"); stopDictation(); }
    );
  }

  function stopDictation() {
    const r = recognizer;
    recognizer = null;
    setRecordingUI(false);
    if (!r) return;
    r.stopContinuousRecognitionAsync(
      () => { r.close(); if (!els.status.classList.contains("error")) setStatus("Diktálás leállítva. Ellenőrizze a szöveget.", "ok"); },
      () => r.close()
    );
  }

  // ---------- Fordítás ----------
  async function translate(to) {
    const text = els.text.value.trim();
    if (!text) { setStatus("Nincs lefordítandó szöveg.", "error"); return; }
    const key = els.trKey.value.trim() || els.speechKey.value.trim();
    const region = els.trRegion.value.trim() || els.speechRegion.value.trim();
    if (!key) { els.settings.open = true; setStatus("Hiányzik a Translator kulcs.", "error"); return; }

    const headers = {
      "Ocp-Apim-Subscription-Key": key,
      "Content-Type": "application/json",
    };
    if (region) headers["Ocp-Apim-Subscription-Region"] = region;

    setStatus(to === "de" ? "Fordítás németre…" : "Fordítás angolra…", "");
    [els.toDe, els.toEn].forEach((b) => (b.disabled = true));
    try {
      const res = await fetch(
        "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to=" + to,
        { method: "POST", headers, body: JSON.stringify([{ Text: text }]) }
      );
      const data = await res.json();
      if (!res.ok) throw new Error((data.error && data.error.message) || res.statusText);
      pushHistory();
      els.text.value = data[0].translations[0].text;
      setStatus("Lefordítva. A „Visszavonás” gombbal visszaállíthatja az eredetit.", "ok");
    } catch (e) {
      setStatus("Fordítási hiba: " + (e.message || e), "error");
    } finally {
      [els.toDe, els.toEn].forEach((b) => (b.disabled = false));
    }
  }

  // ---------- Beszúrás az e-mailbe ----------
  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function insertIntoMail() {
    const text = els.text.value.trim();
    if (!text) { setStatus("Nincs beszúrandó szöveg.", "error"); return; }
    if (!inOutlook() || !Office.context.mailbox.item || !Office.context.mailbox.item.body) {
      setStatus("Ez csak az Outlookban, e-mail írása közben működik.", "error");
      return;
    }
    const body = Office.context.mailbox.item.body;
    body.getTypeAsync((typeRes) => {
      const isHtml = typeRes.status === Office.AsyncResultStatus.Succeeded &&
        typeRes.value === Office.CoercionType.Html;
      const data = isHtml
        ? text.split(/\r?\n/).map((l) => escapeHtml(l) || "&nbsp;").map((l) => "<div>" + l + "</div>").join("")
        : text;
      body.setSelectedDataAsync(
        data,
        { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text },
        (r) => {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            setStatus("Beszúrva az e-mailbe. Ellenőrzés után nyomja meg a Küldés gombot.", "ok");
          } else {
            setStatus("Beszúrási hiba: " + r.error.message, "error");
          }
        }
      );
    });
  }

  // ---------- Eseménykezelők ----------
  function wire() {
    els.micBtn.addEventListener("click", () => (recognizer ? stopDictation() : startDictation()));
    els.toDe.addEventListener("click", () => translate("de"));
    els.toEn.addEventListener("click", () => translate("en"));
    els.undo.addEventListener("click", () => {
      if (!history.length) return;
      els.text.value = history.pop();
      els.undo.disabled = history.length === 0;
      setStatus("Visszaállítva.", "ok");
    });
    els.clear.addEventListener("click", () => {
      if (!els.text.value) return;
      pushHistory();
      els.text.value = "";
      setStatus("", "");
    });
    els.insertBtn.addEventListener("click", insertIntoMail);
    els.saveSettings.addEventListener("click", saveSettings);
    els.lang.addEventListener("change", () => {
      localStorage.setItem("diktalas.lang", els.lang.value);
      if (inOutlook() && Office.context.roamingSettings) {
        Office.context.roamingSettings.set("lang", els.lang.value);
        Office.context.roamingSettings.saveAsync();
      }
    });
  }

  function init() { wire(); loadSettings(); }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(init);
  } else {
    document.addEventListener("DOMContentLoaded", init);
  }
})();
