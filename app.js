/* ============================================================
   Décharge mentale — logique de l'application
   Tout tourne côté client (pas de serveur). Les données restent
   sur le téléphone (localStorage) sauf ce qui est envoyé vers
   Microsoft Graph (Outlook) une fois connecté.
   ============================================================ */

const STORE_KEYS = {
  categories: "dm_categories",
  clientId: "dm_client_id",
  quickMode: "dm_quick_mode",
  journal: "dm_journal",
  msalAccountId: "dm_msal_account_id",
};

const DEFAULT_CATEGORIES = [
  { id: "perso", label: "Perso", color: "preset10", keywords: ["maison", "courses", "enfant", "enfants", "famille"] },
  { id: "entreprise1", label: "Jambes-Machines", color: "preset1", keywords: ["jambes-machines", "jambes machines"] },
  { id: "entreprise2", label: "Roger Bauwens", color: "preset5", keywords: ["roger bauwens", "bauwens"] },
  { id: "entreprise3", label: "Baudouin Vergote", color: "preset20", keywords: ["baudouin vergote", "vergote"] },
];

// Couleurs Outlook "preset" disponibles côté Graph (masterCategories)
const OUTLOOK_PRESETS = [
  "preset0","preset1","preset2","preset3","preset4","preset5","preset6","preset7",
  "preset8","preset9","preset10","preset11","preset12","preset13","preset14",
  "preset15","preset16","preset17","preset18","preset19","preset20","preset21",
  "preset22","preset23","preset24",
];
const PRESET_SWATCH = {
  preset0:"#a6272a", preset1:"#e04f5f", preset2:"#e98124", preset3:"#e3b122", preset4:"#84b301",
  preset5:"#0b6a0b", preset6:"#0f8299", preset7:"#0a53a8", preset8:"#3960b4", preset9:"#7a5ea8",
  preset10:"#c4487a", preset11:"#8c8c8c", preset12:"#7b7c8c", preset13:"#a6272a", preset14:"#e04f5f",
  preset15:"#e98124", preset16:"#e3b122", preset17:"#84b301", preset18:"#0b6a0b", preset19:"#0f8299",
  preset20:"#0a53a8", preset21:"#3960b4", preset22:"#7a5ea8", preset23:"#c4487a", preset24:"#8c8c8c",
};

const SHOPPING_LIST_NAME = "🛒 Courses";
const TASK_LIST_NAME = "🧠 Décharge mentale";
const SHOPPING_KEYWORDS = [
  "acheter", "racheter", "recommander", "commander", "prendre au magasin",
  "il faut du", "il faut de la", "il faut des", "au supermarché", "à la pharmacie",
  "courses", "faire les courses",
];

// Mots-clés qu'on peut dire explicitement pour trancher soi-même (priorité sur la détection auto)
const EXPLICIT_TASK_KEYWORDS = ["tâche", "tache", "todo", "to-do", "à faire", "a faire"];
const EXPLICIT_EVENT_KEYWORDS = ["rendez-vous", "rendez vous", "rdv", "calendrier", "événement", "evenement", "meeting", "réunion", "reunion"];

// Verbes d'action qui, combinés à une date, indiquent plutôt une tâche avec échéance
// qu'un rendez-vous ("rappeler le 5" = tâche, "visite le 5" = événement)
const TASK_ACTION_VERBS = [
  "rappeler", "appeler", "envoyer", "préparer", "preparer", "finir", "terminer",
  "relancer", "commander", "réserver", "reserver", "payer", "contacter", "répondre",
  "repondre", "valider", "vérifier", "verifier", "confirmer", "transmettre",
  "demander", "récupérer", "recuperer", "renvoyer",
];

// ---------- Utilitaires stockage ----------
function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}
function saveJSON(key, val) {
  localStorage.setItem(key, JSON.stringify(val));
}

function getCategories() {
  return loadJSON(STORE_KEYS.categories, DEFAULT_CATEGORIES);
}
function setCategories(cats) {
  saveJSON(STORE_KEYS.categories, cats);
}
function getJournal() {
  return loadJSON(STORE_KEYS.journal, []);
}
function setJournal(items) {
  saveJSON(STORE_KEYS.journal, items);
}
function getClientId() {
  return localStorage.getItem(STORE_KEYS.clientId) || "";
}
function getQuickMode() {
  return localStorage.getItem(STORE_KEYS.quickMode) === "1";
}

// ============================================================
// 1. RECONNAISSANCE VOCALE
// ============================================================
let recognition = null;
let isRecording = false;

function initSpeech() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) return null;
  const r = new SR();
  r.lang = "fr-FR";
  // Le mode "continu" (continuous=true) est peu fiable sur Chrome Android : il redémarre
  // en boucle tout seul et répète des bouts de phrase, surtout avec du bruit de fond.
  // Une phrase par écoute est plus lent mais fiable ; on retouche le micro pour continuer.
  r.continuous = false;
  r.interimResults = true;
  return r;
}

function setupMic() {
  const micBtn = document.getElementById("btn-mic");
  const micHint = document.getElementById("mic-hint");
  const transcript = document.getElementById("live-transcript");

  recognition = initSpeech();
  if (!recognition) {
    micHint.textContent = "Micro non supporté ici — écris ton texte";
    micBtn.disabled = true;
    micBtn.style.opacity = 0.35;
    return;
  }

  let finalText = "";
  let finalizedThisSession = false;

  recognition.onresult = (event) => {
    let interim = "";
    for (let i = event.resultIndex; i < event.results.length; i++) {
      const chunk = event.results[i][0].transcript;
      // continuous=false : normalement un seul résultat final par écoute, mais on se
      // protège quand même d'un éventuel doublon envoyé deux fois par le navigateur.
      if (event.results[i].isFinal) {
        if (!finalizedThisSession) {
          finalText += chunk + " ";
          finalizedThisSession = true;
        }
      } else {
        interim += chunk;
      }
    }
    transcript.textContent = (finalText + interim).trim();
    updateNextButton();
  };

  recognition.onerror = () => {
    resetMicUI();
  };

  // Pas de redémarrage automatique ici : sur Chrome Android, relancer en boucle
  // provoque des répétitions. L'écoute s'arrête simplement ; on retouche le micro
  // pour dicter la suite (le texte déjà transcrit est conservé).
  recognition.onend = () => {
    resetMicUI();
  };

  micBtn.addEventListener("click", () => {
    if (isRecording) {
      try { recognition.stop(); } catch (e) {}
      resetMicUI();
    } else {
      finalText = transcript.textContent ? transcript.textContent + " " : "";
      finalizedThisSession = false;
      startRecording();
    }
  });

  function startRecording() {
    isRecording = true;
    micBtn.classList.add("recording");
    micHint.textContent = "Écoute… touche pour arrêter";
    try { recognition.start(); } catch (e) { resetMicUI(); }
  }

  function resetMicUI() {
    isRecording = false;
    micBtn.classList.remove("recording");
    micHint.textContent = "Touche pour parler";
  }
}

function updateNextButton() {
  const text = document.getElementById("live-transcript").textContent.trim();
  document.getElementById("btn-next").disabled = text.length === 0;
}

// ============================================================
// 2. ANALYSE DU TEXTE : type (tâche / événement / courses) + date
// ============================================================
const WEEKDAYS = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];
// Chaque entrée : [nom(s) reconnus, index du mois 0-11]. Les variantes avec/sans accent
// pointent explicitement vers le bon mois (pas de calcul par modulo, source du bug précédent).
const MONTHS = [
  ["janvier", 0], ["février", 1], ["fevrier", 1], ["mars", 2], ["avril", 3],
  ["mai", 4], ["juin", 5], ["juillet", 6], ["août", 7], ["aout", 7],
  ["septembre", 8], ["octobre", 9], ["novembre", 10], ["décembre", 11], ["decembre", 11],
];

function parseFrenchDateTime(text, now = new Date()) {
  const t = text.toLowerCase();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let date = null;
  let hasTime = false;

  // "dans X jours / heures / minutes / semaines"
  let m = t.match(/dans\s+(\d+)\s*(minute|heure|jour|semaine)s?/);
  if (m) {
    const n = parseInt(m[1], 10);
    date = new Date(now);
    if (m[2] === "minute") date.setMinutes(date.getMinutes() + n);
    else if (m[2] === "heure") { date.setHours(date.getHours() + n); hasTime = true; }
    else if (m[2] === "jour") date.setDate(date.getDate() + n);
    else if (m[2] === "semaine") date.setDate(date.getDate() + n * 7);
    if (m[2] === "minute") hasTime = true;
  }

  // aujourd'hui / demain / après-demain
  if (!date) {
    if (/\baujourd'?hui\b/.test(t)) date = new Date(now);
    else if (/\bapr[eè]s[- ]demain\b/.test(t)) { date = new Date(now); date.setDate(date.getDate() + 2); }
    else if (/\bdemain\b/.test(t)) { date = new Date(now); date.setDate(date.getDate() + 1); }
  }

  // jour de la semaine ("lundi", "lundi prochain")
  if (!date) {
    for (let i = 0; i < WEEKDAYS.length; i++) {
      const re = new RegExp("\\b" + WEEKDAYS[i] + "\\b");
      if (re.test(t)) {
        date = new Date(now);
        let diff = (i - now.getDay() + 7) % 7;
        if (diff === 0) diff = 7; // le prochain, pas aujourd'hui
        date.setDate(date.getDate() + diff);
        break;
      }
    }
  }

  // date explicite "12/09" ou "12-09" ou "12 septembre"
  if (!date) {
    m = t.match(/\b(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?\b/);
    if (m) {
      const day = parseInt(m[1], 10), month = parseInt(m[2], 10) - 1;
      let year = m[3] ? parseInt(m[3], 10) : now.getFullYear();
      if (year < 100) year += 2000;
      date = new Date(year, month, day);
    }
  }
  if (!date) {
    for (const [name, monthIndex] of MONTHS) {
      const re = new RegExp("\\b(\\d{1,2})\\s+" + name + "\\b");
      m = t.match(re);
      if (m) {
        const day = parseInt(m[1], 10);
        date = new Date(now.getFullYear(), monthIndex, day);
        if (date < todayMidnight) date.setFullYear(date.getFullYear() + 1);
        break;
      }
    }
  }

  // heure explicite "14h", "14h30", "à 9h", "9:30"
  const timeMatch = t.match(/\b(\d{1,2})\s*[h:]\s*(\d{2})?\b/);
  if (timeMatch) {
    const hh = parseInt(timeMatch[1], 10);
    const mm = timeMatch[2] ? parseInt(timeMatch[2], 10) : 0;
    if (hh >= 0 && hh <= 23) {
      if (!date) date = new Date(now);
      date.setHours(hh, mm, 0, 0);
      hasTime = true;
    }
  }

  if (!date) return null;
  if (!hasTime) date.setHours(9, 0, 0, 0); // valeur par défaut si seule la date est connue
  return { date, hasTime };
}

function detectShopping(text) {
  const t = text.toLowerCase();
  return SHOPPING_KEYWORDS.some((k) => t.includes(k));
}

function detectCategory(text, categories) {
  const t = text.toLowerCase();
  const matches = categories.filter((c) =>
    (c.keywords || []).some((k) => k.trim() && t.includes(k.trim().toLowerCase()))
  );
  if (matches.length === 1) return matches[0].id;
  return null; // ambigu ou aucun -> l'utilisateur choisit
}

function analyzeText(text) {
  const categories = getCategories();
  const t = text.toLowerCase();
  const shopping = detectShopping(text);
  const dateInfo = parseFrenchDateTime(text);
  const categoryId = detectCategory(text, categories) || categories[0]?.id || null;

  const explicitTask = EXPLICIT_TASK_KEYWORDS.some((k) => t.includes(k));
  const explicitEvent = EXPLICIT_EVENT_KEYWORDS.some((k) => t.includes(k));
  const hasTaskVerb = TASK_ACTION_VERBS.some((v) => t.includes(v));

  let type = "task";
  if (explicitTask) {
    type = "task";
  } else if (explicitEvent) {
    type = "event";
  } else if (shopping) {
    type = "shopping";
  } else if (dateInfo) {
    // une date détectée => plutôt un événement (visite, rdv...), sauf si le texte
    // contient un verbe d'action typique d'une tâche ("rappeler le 5" = tâche)
    type = hasTaskVerb ? "task" : "event";
  }

  return { type, categoryId, dateInfo };
}

// ============================================================
// 3. MSAL / MICROSOFT GRAPH
// ============================================================
let msalInstance = null;
let currentAccount = null;

const GRAPH_SCOPES = ["User.Read", "Calendars.ReadWrite", "Tasks.ReadWrite"];

function getRedirectUri() {
  return window.location.origin + window.location.pathname;
}

async function initMsal() {
  const clientId = getClientId();
  if (!clientId || !window.msal) return false;

  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: getRedirectUri(),
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  });

  await msalInstance.initialize();

  try {
    const resp = await msalInstance.handleRedirectPromise();
    if (resp && resp.account) currentAccount = resp.account;
  } catch (e) {
    console.warn("Erreur redirect MSAL", e);
  }

  if (!currentAccount) {
    const savedId = localStorage.getItem(STORE_KEYS.msalAccountId);
    const accounts = msalInstance.getAllAccounts();
    currentAccount = accounts.find((a) => a.homeAccountId === savedId) || accounts[0] || null;
  }

  updateAuthUI();
  return true;
}

async function login() {
  if (!msalInstance) {
    if (!getClientId()) {
      alert("Renseigne d'abord ton Client ID Azure dans Réglages.");
    } else if (!window.msal) {
      alert(
        "La bibliothèque de connexion Microsoft n'a pas pu se charger (réseau/pare-feu ?). " +
        "Vérifie ta connexion internet et recharge la page. Si ça persiste sur un réseau " +
        "professionnel, il faudra peut-être demander à ton service informatique de débloquer " +
        "cdn.jsdelivr.net."
      );
    } else {
      alert("La connexion n'est pas encore prête, réessaie dans un instant.");
    }
    openScreen("settings");
    return;
  }
  try {
    await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
  } catch (e) {
    alert("Connexion impossible : " + e.message);
  }
}

function logout() {
  if (!msalInstance || !currentAccount) return;
  localStorage.removeItem(STORE_KEYS.msalAccountId);
  msalInstance.logoutRedirect({ account: currentAccount });
}

async function getGraphToken(interactive = false) {
  if (!msalInstance || !currentAccount) return null;
  try {
    const resp = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: currentAccount });
    return resp.accessToken;
  } catch (e) {
    // Ne JAMAIS rediriger automatiquement en arrière-plan (retry offline, chargement de
    // la page...) : ça enverrait l'utilisateur sur l'écran de connexion Microsoft sans
    // qu'il ait rien demandé. On ne redirige que si l'appel vient d'une action explicite
    // (ex: bouton "Envoyer vers Outlook").
    if (interactive) {
      try {
        await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account: currentAccount });
      } catch (e2) {
        console.warn(e2);
      }
    }
    return null;
  }
}

function updateAuthUI() {
  const connected = !!currentAccount;
  document.getElementById("auth-banner").classList.toggle("hidden", connected);
  const statusEl = document.getElementById("account-status");
  if (statusEl) {
    statusEl.textContent = connected
      ? "Connecté : " + (currentAccount.username || currentAccount.name)
      : "Non connecté";
  }
  if (connected) localStorage.setItem(STORE_KEYS.msalAccountId, currentAccount.homeAccountId);
}

async function graphFetch(path, options = {}, interactive = false) {
  const token = await getGraphToken(interactive);
  if (!token) throw new Error("no-auth");
  const resp = await fetch("https://graph.microsoft.com/v1.0" + path, {
    ...options,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!resp.ok) {
    const errText = await resp.text().catch(() => "");
    throw new Error("graph-error " + resp.status + " " + errText);
  }
  if (resp.status === 204) return null;
  return resp.json();
}

let cachedMasterCategories = null;
async function ensureOutlookCategory(name, interactive = false) {
  if (!cachedMasterCategories) {
    const list = await graphFetch("/me/outlook/masterCategories", {}, interactive);
    cachedMasterCategories = list.value || [];
  }
  let existing = cachedMasterCategories.find((c) => c.displayName === name);
  if (existing) return existing;
  const usedColors = new Set(cachedMasterCategories.map((c) => c.color));
  const color = OUTLOOK_PRESETS.find((p) => !usedColors.has(p)) || "preset11";
  const created = await graphFetch("/me/outlook/masterCategories", {
    method: "POST",
    body: JSON.stringify({ displayName: name, color }),
  }, interactive);
  cachedMasterCategories.push(created);
  return created;
}

let cachedTaskListId = null;
async function ensureTaskList(interactive = false) {
  if (cachedTaskListId) return cachedTaskListId;
  const lists = await graphFetch("/me/todo/lists", {}, interactive);
  let list = (lists.value || []).find((l) => l.displayName === TASK_LIST_NAME);
  if (!list) {
    list = await graphFetch("/me/todo/lists", {
      method: "POST",
      body: JSON.stringify({ displayName: TASK_LIST_NAME }),
    }, interactive);
  }
  cachedTaskListId = list.id;
  return cachedTaskListId;
}

function localISOString(date) {
  // Graph veut une date/heure "flottante" + un fuseau explicite
  const pad = (n) => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}:00`;
}
function tz() {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || "Europe/Brussels";
}

async function sendItemToOutlook(item, interactive = false) {
  const categories = getCategories();
  const cat = categories.find((c) => c.id === item.categoryId);
  const catLabel = cat ? cat.label : null;
  const categoryLabels = [];
  if (catLabel) {
    await ensureOutlookCategory(catLabel, interactive);
    categoryLabels.push(catLabel);
  }
  if (item.type === "shopping") {
    await ensureOutlookCategory(SHOPPING_LIST_NAME, interactive);
    categoryLabels.push(SHOPPING_LIST_NAME);
  }

  if (item.type === "event") {
    const start = new Date(item.dateISO);
    const isAllDay = !item.hasTime;
    let eventBody;
    if (isAllDay) {
      // Événement "jour entier" : Microsoft Graph attend le début/la fin en UTC pur
      // (minuit UTC) pour les événements isAllDay. Passer le fuseau local ici est un
      // piège connu de l'API : l'événement peut alors se décaler d'un jour dans Outlook.
      const pad = (n) => String(n).padStart(2, "0");
      const y = start.getFullYear(), mo = start.getMonth(), d = start.getDate();
      const startStr = `${y}-${pad(mo + 1)}-${pad(d)}T00:00:00.000`;
      const nextDay = new Date(y, mo, d + 1);
      const endStr = `${nextDay.getFullYear()}-${pad(nextDay.getMonth() + 1)}-${pad(nextDay.getDate())}T00:00:00.000`;
      eventBody = {
        subject: item.text.slice(0, 120),
        body: { contentType: "text", content: item.text },
        isAllDay: true,
        start: { dateTime: startStr, timeZone: "UTC" },
        end: { dateTime: endStr, timeZone: "UTC" },
        categories: categoryLabels,
      };
    } else {
      const end = new Date(start.getTime() + 60 * 60 * 1000);
      eventBody = {
        subject: item.text.slice(0, 120),
        body: { contentType: "text", content: item.text },
        start: { dateTime: localISOString(start), timeZone: tz() },
        end: { dateTime: localISOString(end), timeZone: tz() },
        categories: categoryLabels,
      };
    }
    await graphFetch("/me/events", {
      method: "POST",
      body: JSON.stringify(eventBody),
    }, interactive);
  } else {
    const listId = await ensureTaskList(interactive);
    const body = {
      title: item.text.slice(0, 250),
      body: { content: item.text, contentType: "text" },
      categories: categoryLabels,
    };
    if (item.dateISO) {
      const d = new Date(item.dateISO);
      body.dueDateTime = { dateTime: localISOString(d), timeZone: tz() };
    }
    await graphFetch(`/me/todo/lists/${listId}/tasks`, {
      method: "POST",
      body: JSON.stringify(body),
    }, interactive);
  }
}

// ============================================================
// 4. JOURNAL / FILE D'ATTENTE (fonctionne hors-ligne)
// ============================================================
function addToJournal(item) {
  const journal = getJournal();
  journal.unshift(item);
  setJournal(journal.slice(0, 200));
  renderJournalBadge();
}
function updateJournalItem(id, patch) {
  const journal = getJournal();
  const idx = journal.findIndex((i) => i.id === id);
  if (idx >= 0) {
    journal[idx] = { ...journal[idx], ...patch };
    setJournal(journal);
  }
  renderJournalBadge();
}
function renderJournalBadge() {
  const pending = getJournal().filter((i) => i.status !== "sent").length;
  const badge = document.getElementById("pending-badge");
  badge.textContent = pending;
  badge.classList.toggle("hidden", pending === 0);
}

async function trySend(item, interactive = false) {
  try {
    await sendItemToOutlook(item, interactive);
    updateJournalItem(item.id, { status: "sent", error: null });
    return true;
  } catch (e) {
    const msg = e.message === "no-auth" ? "Connexion Microsoft requise" : "Erreur d'envoi";
    updateJournalItem(item.id, { status: "error", error: msg });
    return false;
  }
}

async function retryAllPending() {
  const journal = getJournal().filter((i) => i.status !== "sent");
  for (const item of journal) {
    await trySend(item);
  }
  renderJournal();
}

window.addEventListener("online", () => retryAllPending());

// ============================================================
// 5. NAVIGATION / UI
// ============================================================
let pendingItem = null; // item en cours de confirmation

function openScreen(name) {
  document.querySelectorAll(".screen").forEach((s) => s.classList.add("hidden"));
  document.getElementById("screen-" + name).classList.remove("hidden");
  if (name === "journal") renderJournal();
  if (name === "settings") renderSettings();
}

function renderQuickCats() {
  const wrap = document.getElementById("quick-cats");
  wrap.innerHTML = "";
  getCategories().forEach((c) => {
    const chip = document.createElement("span");
    chip.className = "quick-cat-chip";
    chip.textContent = c.label;
    wrap.appendChild(chip);
  });
}

function goToConfirm() {
  const text = document.getElementById("live-transcript").textContent.trim();
  if (!text) return;
  const analysis = analyzeText(text);
  pendingItem = {
    id: "it_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7),
    text,
    type: analysis.type,
    categoryId: analysis.categoryId,
    dateISO: analysis.dateInfo ? analysis.dateInfo.date.toISOString() : null,
    hasTime: analysis.dateInfo ? analysis.dateInfo.hasTime : false,
    createdAt: new Date().toISOString(),
    status: "draft",
  };
  renderConfirmScreen();
  openScreen("confirm");
  startQuickModeTimer();
}

// ---------- Mode rapide : envoi automatique après quelques secondes d'inactivité ----------
let quickModeTimeout = null;
let quickModeInterval = null;

function clearQuickModeTimer() {
  clearTimeout(quickModeTimeout);
  clearInterval(quickModeInterval);
  quickModeTimeout = null;
  quickModeInterval = null;
  const statusEl = document.getElementById("save-status");
  if (statusEl && statusEl.dataset.quickmode === "1") {
    statusEl.textContent = "";
    delete statusEl.dataset.quickmode;
  }
}

function startQuickModeTimer() {
  if (!getQuickMode()) return;
  clearQuickModeTimer();
  let secondsLeft = 4;
  const statusEl = document.getElementById("save-status");
  const render = () => {
    statusEl.textContent = `Envoi automatique dans ${secondsLeft}s… touche l'écran pour corriger`;
    statusEl.dataset.quickmode = "1";
  };
  render();
  quickModeInterval = setInterval(() => {
    secondsLeft -= 1;
    if (secondsLeft > 0) render();
  }, 1000);
  quickModeTimeout = setTimeout(() => {
    clearInterval(quickModeInterval);
    quickModeInterval = null;
    document.getElementById("btn-save").click();
  }, 4000);
}

function renderConfirmScreen() {
  document.getElementById("confirm-text").value = pendingItem.text;

  const catRow = document.getElementById("confirm-cats");
  catRow.innerHTML = "";
  getCategories().forEach((c) => {
    const btn = document.createElement("button");
    btn.className = "chip" + (c.id === pendingItem.categoryId ? " selected" : "");
    btn.textContent = c.label;
    btn.addEventListener("click", () => {
      pendingItem.categoryId = c.id;
      renderConfirmScreen();
    });
    catRow.appendChild(btn);
  });

  document.querySelectorAll("#screen-confirm .chip-row .chip[data-type]").forEach((btn) => {
    btn.classList.toggle("selected", btn.dataset.type === pendingItem.type);
    btn.onclick = () => {
      pendingItem.type = btn.dataset.type;
      renderConfirmScreen();
    };
  });

  const dateField = document.getElementById("date-field");
  const dtInput = document.getElementById("confirm-datetime");
  const alldayBox = document.getElementById("confirm-allday");
  dateField.style.display = pendingItem.type === "shopping" ? "none" : "flex";
  document.getElementById("date-label").textContent = pendingItem.type === "event" ? "Quand" : "Échéance (optionnel)";

  if (pendingItem.dateISO) {
    const d = new Date(pendingItem.dateISO);
    const pad = (n) => String(n).padStart(2, "0");
    dtInput.value = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  } else {
    dtInput.value = "";
  }
  alldayBox.checked = !pendingItem.hasTime;
}

function bindConfirmScreen() {
  // toute interaction sur l'écran annule le compte à rebours du mode rapide
  document.getElementById("screen-confirm").addEventListener("pointerdown", (e) => {
    if (e.target.closest("#btn-save")) return; // laisse le clic auto sur "Envoyer" passer
    clearQuickModeTimer();
  });

  document.getElementById("confirm-text").addEventListener("input", (e) => {
    pendingItem.text = e.target.value;
  });
  document.getElementById("confirm-datetime").addEventListener("input", (e) => {
    if (e.target.value) {
      pendingItem.dateISO = new Date(e.target.value).toISOString();
      pendingItem.hasTime = true;
      document.getElementById("confirm-allday").checked = false;
    } else {
      pendingItem.dateISO = null;
    }
  });
  document.getElementById("confirm-allday").addEventListener("change", (e) => {
    pendingItem.hasTime = !e.target.checked;
  });

  document.getElementById("btn-save").addEventListener("click", async () => {
    clearQuickModeTimer();
    const statusEl = document.getElementById("save-status");
    statusEl.textContent = "Envoi en cours…";
    pendingItem.status = "pending";
    addToJournal(pendingItem);
    const ok = await trySend(pendingItem, true);
    statusEl.textContent = ok ? "Envoyé vers Outlook ✅" : "Pas encore envoyé — retenté automatiquement (voir Journal)";
    setTimeout(() => {
      document.getElementById("live-transcript").textContent = "";
      updateNextButton();
      openScreen("capture");
    }, 700);
  });

  document.getElementById("btn-back-confirm").addEventListener("click", () => {
    clearQuickModeTimer();
    openScreen("capture");
  });
}

function renderJournal() {
  const list = document.getElementById("journal-list");
  const items = getJournal();
  if (items.length === 0) {
    list.innerHTML = '<div class="journal-empty">Rien pour l\'instant. Ta tête est libre 🧘</div>';
    return;
  }
  const categories = getCategories();
  list.innerHTML = "";
  items.forEach((item) => {
    const cat = categories.find((c) => c.id === item.categoryId);
    const div = document.createElement("div");
    div.className = "journal-item";
    const statusClass = item.status === "sent" ? "status-sent" : item.status === "error" ? "status-error" : "status-pending";
    const statusText = item.status === "sent" ? "Envoyé" : item.status === "error" ? "Erreur" : "En attente";
    const typeIcon = item.type === "event" ? "📅" : item.type === "shopping" ? "🛒" : "✅";
    div.innerHTML = `
      <div class="jtext">${typeIcon} ${escapeHtml(item.text)}</div>
      <div class="jmeta">
        <span>${cat ? escapeHtml(cat.label) : "—"} · ${new Date(item.createdAt).toLocaleString("fr-FR")}</span>
        <span class="status-pill ${statusClass}">${statusText}</span>
      </div>
    `;
    if (item.status === "error") {
      const retryBtn = document.createElement("button");
      retryBtn.className = "btn-secondary";
      retryBtn.textContent = "Réessayer";
      retryBtn.style.alignSelf = "flex-start";
      retryBtn.addEventListener("click", async () => {
        await trySend(item, true);
        renderJournal();
      });
      div.appendChild(retryBtn);
    }
    list.appendChild(div);
  });
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

// ---------- Réglages ----------
function renderSettings() {
  document.getElementById("client-id-input").value = getClientId();
  document.getElementById("quick-mode").checked = getQuickMode();
  updateAuthUI();
  renderCategoriesEditor();
}

function renderCategoriesEditor() {
  const wrap = document.getElementById("categories-editor");
  wrap.innerHTML = "";
  const categories = getCategories();
  categories.forEach((cat, idx) => {
    const row = document.createElement("div");
    row.className = "category-row";
    row.innerHTML = `
      <div class="row-top">
        <span class="swatch" style="background:${PRESET_SWATCH[cat.color] || "#888"}"></span>
        <input type="text" class="cat-label" value="${escapeHtml(cat.label)}" placeholder="Nom de la catégorie" />
        <button class="remove-cat" aria-label="Supprimer">✕</button>
      </div>
      <input type="text" class="cat-keywords" value="${escapeHtml((cat.keywords || []).join(", "))}" placeholder="mots-clés, séparés, par virgule" />
    `;
    row.querySelector(".cat-label").addEventListener("input", (e) => {
      categories[idx].label = e.target.value;
    });
    row.querySelector(".cat-keywords").addEventListener("input", (e) => {
      categories[idx].keywords = e.target.value.split(",").map((s) => s.trim()).filter(Boolean);
    });
    row.querySelector(".remove-cat").addEventListener("click", () => {
      categories.splice(idx, 1);
      setCategories(categories);
      renderCategoriesEditor();
    });
    wrap.appendChild(row);
  });
  window.__editingCategories = categories;
}

function bindSettingsScreen() {
  document.getElementById("btn-back-settings").addEventListener("click", () => openScreen("capture"));
  document.getElementById("btn-login2").addEventListener("click", login);
  document.getElementById("btn-logout").addEventListener("click", logout);

  document.getElementById("btn-add-category").addEventListener("click", () => {
    const categories = window.__editingCategories || getCategories();
    const usedColors = new Set(categories.map((c) => c.color));
    const color = OUTLOOK_PRESETS.find((p) => !usedColors.has(p)) || "preset11";
    categories.push({ id: "cat_" + Date.now(), label: "Nouvelle catégorie", color, keywords: [] });
    setCategories(categories);
    renderCategoriesEditor();
  });

  document.getElementById("btn-save-settings").addEventListener("click", async () => {
    const clientId = document.getElementById("client-id-input").value.trim();
    const prevClientId = getClientId();
    localStorage.setItem(STORE_KEYS.clientId, clientId);
    localStorage.setItem(STORE_KEYS.quickMode, document.getElementById("quick-mode").checked ? "1" : "0");
    setCategories(window.__editingCategories || getCategories());
    renderQuickCats();

    if (clientId && clientId !== prevClientId) {
      await initMsal();
    }
    alert("Réglages enregistrés ✅");
    openScreen("capture");
  });
}

// ============================================================
// 6. INITIALISATION
// ============================================================
async function init() {
  if (!localStorage.getItem(STORE_KEYS.categories)) {
    setCategories(DEFAULT_CATEGORIES);
  }

  setupMic();
  renderQuickCats();
  bindConfirmScreen();
  bindSettingsScreen();
  renderJournalBadge();

  // saisie manuelle (clavier) dans la zone de transcription : active/désactive "Continuer"
  document.getElementById("live-transcript").addEventListener("input", updateNextButton);

  document.getElementById("btn-next").addEventListener("click", goToConfirm);
  document.getElementById("btn-settings").addEventListener("click", () => openScreen("settings"));
  document.getElementById("btn-journal").addEventListener("click", () => openScreen("journal"));
  document.getElementById("btn-back-journal").addEventListener("click", () => openScreen("capture"));
  document.getElementById("btn-login").addEventListener("click", login);
  document.getElementById("btn-retry-all").addEventListener("click", retryAllPending);

  if (getClientId()) {
    await initMsal();
  }

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }

  // tentative d'envoi des éléments en attente au démarrage
  retryAllPending();
}

document.addEventListener("DOMContentLoaded", init);
