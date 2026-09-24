/**
 * Gemini AI Assistant for Excel - Taskpane Controller
 * 100% Free, Client-side Bring-Your-Own-Key (BYOK) via Google AI Studio
 * Features:
 * - Session Transcript Preservation across Excel reloads
 * - 1-Click Export to Markdown (.md)
 * - 1-Click Save to Excel Worksheet ("AI_Transcript")
 * - Auto Smart Fallback across Gemini 3.8, 3.7, 3.6, 3.5 Flash Lite & 3.1 Pro
 */

// Available Models Definition
const MODEL_NAMES = {
  "auto": "Auto (Smart Fallback)",
  "gemini-3.8-flash": "Gemini 3.8 Flash",
  "gemini-3.7-flash": "Gemini 3.7 Flash",
  "gemini-3.6-flash": "Gemini 3.6 Flash",
  "gemini-3.5-flash-lite": "Gemini 3.5 Flash Lite",
  "gemini-3.1-pro-preview": "Gemini 3.1 Pro (Preview)"
};

const ORDERED_MODELS = [
  "gemini-3.8-flash",
  "gemini-3.7-flash",
  "gemini-3.6-flash",
  "gemini-3.5-flash-lite"
];

function getModelDisplayName(id) {
  return MODEL_NAMES[id] || id;
}

// State
let currentSelection = {
  address: "None",
  rowCount: 0,
  columnCount: 0,
  values: [],
  formulas: [],
  hasData: false
};

let chatHistory = [];
let savedMessages = [];
let isGenerating = false;
let isOfficeReady = false;
let currentApiKey = "";

const DEFAULT_SYSTEM_PROMPT = `You are an elite Excel, Financial Modeling, and Data Analysis AI assistant powered by Google Gemini.
Your job is to assist the user directly within Microsoft Excel.
When suggesting formulas:
1. Always prefer modern, efficient dynamic array formulas (e.g., XLOOKUP, LET, FILTER, SORT, UNIQUE, SUMIFS, LAMBDA) over deprecated legacy formulas (like VLOOKUP or nested IFs).
2. Clearly format formulas inside markdown code blocks with language \`excel\` (e.g., \`\`\`excel\n=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B)\n\`\`\`).
3. Always include the leading '=' sign in formula code blocks so the user can insert it directly with 1 click.
4. If the user provides a data range, examine headers, data types, errors (#N/A, #VALUE!), and give concise, high-impact insights.
5. Keep explanations concise, practical, and directly applicable.`;

// DOM Elements
const chatMessages = document.getElementById("chatMessages");
const promptInput = document.getElementById("promptInput");
const btnSend = document.getElementById("btnSend");
const modelSelector = document.getElementById("modelSelector");
const statusDot = document.getElementById("statusDot");
const selectionBadge = document.getElementById("selectionBadge");
const chkIncludeSelection = document.getElementById("chkIncludeSelection");
const btnRefreshSelection = document.getElementById("btnRefreshSelection");
const settingsModal = document.getElementById("settingsModal");
const btnOpenSettings = document.getElementById("btnOpenSettings");
const btnCloseSettings = document.getElementById("btnCloseSettings");
const btnSaveSettings = document.getElementById("btnSaveSettings");
const btnTestKey = document.getElementById("btnTestKey");
const apiKeyInput = document.getElementById("apiKeyInput");
const systemPromptInput = document.getElementById("systemPromptInput");
const tempSlider = document.getElementById("tempSlider");
const tempValue = document.getElementById("tempValue");
const toggleKeyVisibility = document.getElementById("toggleKeyVisibility");
const testStatus = document.getElementById("testStatus");
const btnClearChat = document.getElementById("btnClearChat");
const btnExportTranscript = document.getElementById("btnExportTranscript");
const btnExportSheet = document.getElementById("btnExportSheet");

// Multi-tier key retrieval
function getStoredApiKey() {
  if (currentApiKey && currentApiKey.length > 5) return currentApiKey;
  
  let key = "";
  try { key = localStorage.getItem("gemini_api_key") || ""; } catch (e) {}
  if (key && key.trim().length > 5) {
    currentApiKey = key.trim();
    return currentApiKey;
  }

  try { key = sessionStorage.getItem("gemini_api_key") || ""; } catch (e) {}
  if (key && key.trim().length > 5) {
    currentApiKey = key.trim();
    return currentApiKey;
  }

  if (apiKeyInput && apiKeyInput.value.trim().length > 5) {
    currentApiKey = apiKeyInput.value.trim();
    return currentApiKey;
  }

  return "";
}

function persistApiKey(key) {
  if (!key) return;
  currentApiKey = key.trim();
  try { localStorage.setItem("gemini_api_key", currentApiKey); } catch (e) {}
  try { sessionStorage.setItem("gemini_api_key", currentApiKey); } catch (e) {}

  if (isOfficeReady && Office.context?.document?.settings) {
    try {
      Office.context.document.settings.set("gemini_api_key", currentApiKey);
      Office.context.document.settings.saveAsync();
    } catch (e) {}
  }
}

// Restore saved transcript from localStorage
function restoreChatTranscript() {
  try {
    const raw = localStorage.getItem("gemini_chat_transcript");
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        savedMessages = parsed;
        const welcomeCard = document.querySelector(".welcome-card");
        if (welcomeCard) welcomeCard.remove();

        savedMessages.forEach((msg) => {
          renderMessageNode(msg.role, msg.content, msg.time, msg.selection, msg.modelBadge);
          chatHistory.push({
            role: msg.role === "assistant" ? "model" : "user",
            parts: [{ text: msg.content }]
          });
        });
        console.log(`Restored ${savedMessages.length} messages from previous session.`);
      }
    }
  } catch (err) {
    console.warn("Could not restore chat transcript:", err);
  }
}

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    isOfficeReady = true;
    console.log("Running inside Microsoft Excel");
    
    // Check if document settings has the key stored
    if (Office.context?.document?.settings) {
      try {
        const docKey = Office.context.document.settings.get("gemini_api_key");
        if (docKey && !currentApiKey) {
          persistApiKey(docKey);
          if (apiKeyInput) apiKeyInput.value = docKey;
          checkApiStatus();
        }
      } catch (e) {}
    }

    // Listen for selection changes in real-time
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => refreshSelection()
    );

    // Initial read
    refreshSelection();
  } else {
    console.log("Running in standalone / browser mode");
    selectionBadge.textContent = "Standalone Mode (Mock A1:C5)";
    currentSelection = {
      address: "Sheet1!A1:C5",
      rowCount: 5,
      columnCount: 3,
      values: [
        ["Date", "Revenue", "Cost"],
        ["2026-01-01", 12500, 8200],
        ["2026-02-01", 14300, 9100],
        ["2026-03-01", 18900, 11400],
        ["2026-04-01", 21200, 12800]
      ],
      formulas: [["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""]],
      hasData: true
    };
  }
});

// App Lifecycle
document.addEventListener("DOMContentLoaded", () => {
  loadSettings();
  setupEventListeners();
  checkApiStatus();
  restoreChatTranscript();
});

function loadSettings() {
  const apiKey = getStoredApiKey();
  let model = "auto";
  try { model = localStorage.getItem("gemini_model") || "auto"; } catch (e) {}

  // Auto-migrate legacy/deprecated models to auto
  if (model === "gemini-2.5-flash" || model === "gemini-2.5-pro" || model === "gemini-2.0-flash") {
    model = "auto";
    try { localStorage.setItem("gemini_model", model); } catch (e) {}
  }

  let sysPrompt = DEFAULT_SYSTEM_PROMPT;
  try { sysPrompt = localStorage.getItem("gemini_system_prompt") || DEFAULT_SYSTEM_PROMPT; } catch (e) {}

  let temp = "0.2";
  try { temp = localStorage.getItem("gemini_temp") || "0.2"; } catch (e) {}

  apiKeyInput.value = apiKey;
  modelSelector.value = model;
  systemPromptInput.value = sysPrompt;
  tempSlider.value = temp;
  tempValue.textContent = temp;
}

function saveSettings() {
  const key = apiKeyInput.value.trim();
  const model = modelSelector.value;
  const sysPrompt = systemPromptInput.value.trim() || DEFAULT_SYSTEM_PROMPT;
  const temp = tempSlider.value;

  persistApiKey(key);

  try {
    localStorage.setItem("gemini_model", model);
    localStorage.setItem("gemini_system_prompt", sysPrompt);
    localStorage.setItem("gemini_temp", temp);
  } catch (e) {
    console.warn("Storage write error:", e);
  }

  checkApiStatus();
  settingsModal.classList.remove("active");
}

function checkApiStatus() {
  const key = getStoredApiKey();
  if (key && key.length > 5) {
    statusDot.className = "status-dot connected";
    statusDot.title = "Gemini API Key configured";
  } else {
    statusDot.className = "status-dot error";
    statusDot.title = "Gemini API Key required";
  }
}

function setupEventListeners() {
  // Model selector change
  modelSelector.addEventListener("change", (e) => {
    try { localStorage.setItem("gemini_model", e.target.value); } catch (e) {}
  });

  // Settings Modal
  btnOpenSettings.addEventListener("click", () => {
    loadSettings();
    testStatus.textContent = "";
    settingsModal.classList.add("active");
  });

  btnCloseSettings.addEventListener("click", () => {
    settingsModal.classList.remove("active");
  });

  settingsModal.addEventListener("click", (e) => {
    if (e.target === settingsModal) settingsModal.classList.remove("active");
  });

  btnSaveSettings.addEventListener("click", saveSettings);

  tempSlider.addEventListener("input", (e) => {
    tempValue.textContent = e.target.value;
  });

  toggleKeyVisibility.addEventListener("click", () => {
    if (apiKeyInput.type === "password") {
      apiKeyInput.type = "text";
      toggleKeyVisibility.textContent = "🙈";
    } else {
      apiKeyInput.type = "password";
      toggleKeyVisibility.textContent = "👁️";
    }
  });

  // Auto-persist immediately when user pastes or types their key
  apiKeyInput.addEventListener("input", () => {
    const key = apiKeyInput.value.trim();
    if (key.length > 5) {
      persistApiKey(key);
      checkApiStatus();
    }
  });

  btnTestKey.addEventListener("click", testApiKeyConnection);

  // Selection Refresh
  btnRefreshSelection.addEventListener("click", () => refreshSelection(true));

  // Chat Actions
  btnSend.addEventListener("click", sendMessage);

  promptInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  promptInput.addEventListener("input", () => {
    promptInput.style.height = "auto";
    promptInput.style.height = Math.min(promptInput.scrollHeight, 120) + "px";
    btnSend.disabled = promptInput.value.trim().length === 0 || isGenerating;
  });

  // Clear Chat History
  btnClearChat.addEventListener("click", () => {
    if (confirm("Clear conversation history and stored transcript?")) {
      chatHistory = [];
      savedMessages = [];
      try { localStorage.removeItem("gemini_chat_transcript"); } catch (e) {}
      chatMessages.innerHTML = "";
      renderWelcomeCard();
    }
  });

  // Export Transcript to Markdown file
  if (btnExportTranscript) {
    btnExportTranscript.addEventListener("click", downloadTranscriptMd);
  }

  // Save Transcript to Excel Worksheet
  if (btnExportSheet) {
    btnExportSheet.addEventListener("click", exportTranscriptToSheet);
  }

  // Quick Action Chips
  document.querySelectorAll(".chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const action = chip.dataset.action;
      handleQuickAction(action);
    });
  });
}

// Read current selection from Excel
async function refreshSelection(showFeedback = false) {
  if (!isOfficeReady) return;

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount", "values", "formulas"]);
      await context.sync();

      currentSelection = {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        hasData: range.rowCount > 0 && range.columnCount > 0
      };

      selectionBadge.textContent = `${range.address} (${range.rowCount}×${range.columnCount})`;
      if (showFeedback) {
        selectionBadge.style.borderColor = "var(--gemini-blue)";
        setTimeout(() => (selectionBadge.style.borderColor = "var(--border)"), 800);
      }
    });
  } catch (err) {
    console.error("Error reading selection:", err);
    selectionBadge.textContent = "Selection: Unavailable";
  }
}

// Quick Actions Handler
function handleQuickAction(action) {
  let prompt = "";
  switch (action) {
    case "fix":
      prompt = "Analyze the selected formula or data. If there is a formula error (#N/A, #VALUE!, #REF!, incorrect logic), explain why and provide the exact corrected formula.";
      break;
    case "formula":
      prompt = "Suggest the most optimal modern Excel formula (XLOOKUP, LET, FILTER, SUMIFS, etc.) to perform the calculation or lookup needed for this data.";
      break;
    case "analyze":
      prompt = "Analyze the selected dataset: provide key summary metrics, notable patterns, trends, potential outliers, and actionable insights.";
      break;
    case "clean":
      prompt = "Examine this data for inconsistencies, trailing spaces, duplicate values, mixed formats, or missing entries. Suggest formula or step-by-step fixes to clean it.";
      break;
    case "explain":
      prompt = "Explain in plain English how the formula in the active cell works, breaking down each argument and function step-by-step.";
      break;
    default:
      prompt = action;
  }

  promptInput.value = prompt;
  promptInput.style.height = "auto";
  promptInput.style.height = Math.min(promptInput.scrollHeight, 120) + "px";
  btnSend.disabled = false;
  sendMessage();
}

// Send Message with Auto Fallback & Demand Spike Handling
async function sendMessage() {
  const text = promptInput.value.trim();
  if (!text || isGenerating) return;

  const apiKey = getStoredApiKey();
  if (!apiKey || apiKey.length < 5) {
    settingsModal.classList.add("active");
    testStatus.innerHTML = "<span style='color:#ef4444;'>Please paste your Google AI Studio API Key above and click Save & Apply.</span>";
    return;
  }

  // Remove welcome card if present
  const welcomeCard = document.querySelector(".welcome-card");
  if (welcomeCard) welcomeCard.remove();

  // Add User Message to UI & history
  appendMessage("user", text, currentSelection.hasData && chkIncludeSelection.checked ? currentSelection.address : null);
  promptInput.value = "";
  promptInput.style.height = "24px";
  btnSend.disabled = true;
  isGenerating = true;

  // Add Assistant Loading Bubble
  const loadingId = "loading-" + Date.now();
  appendLoadingBubble(loadingId);

  // Prepare Prompt Context
  let contextualPrompt = text;
  if (chkIncludeSelection.checked && currentSelection.hasData) {
    const safeValues = (currentSelection.values || []).slice(0, 50);
    const safeFormulas = (currentSelection.formulas || []).slice(0, 50);
    const selectionContext = `
[Active Selection Context in Excel:
- Address: ${currentSelection.address}
- Dimensions: ${currentSelection.rowCount} rows × ${currentSelection.columnCount} columns
- Cell Values: ${JSON.stringify(safeValues)}
- Cell Formulas: ${JSON.stringify(safeFormulas)}
]`;
    contextualPrompt = `${text}\n\n${selectionContext}`;
  }

  try {
    const selectedSetting = modelSelector.value || "auto";
    let sysPrompt = DEFAULT_SYSTEM_PROMPT;
    try { sysPrompt = localStorage.getItem("gemini_system_prompt") || DEFAULT_SYSTEM_PROMPT; } catch (e) {}

    let temp = 0.2;
    try { temp = parseFloat(localStorage.getItem("gemini_temp") || "0.2"); } catch (e) {}

    // Build History
    chatHistory.push({ role: "user", parts: [{ text: contextualPrompt }] });

    const payload = {
      contents: chatHistory,
      systemInstruction: {
        parts: [{ text: sysPrompt }]
      },
      generationConfig: {
        temperature: temp,
        maxOutputTokens: 3000
      }
    };

    // Determine cascade order:
    let candidatesToTry = [];
    if (selectedSetting === "auto") {
      candidatesToTry = [...ORDERED_MODELS];
    } else {
      candidatesToTry = [selectedSetting, ...ORDERED_MODELS.filter(m => m !== selectedSetting)];
    }

    let finalReply = null;
    let successfulModel = null;
    let lastError = "";

    for (let i = 0; i < candidatesToTry.length; i++) {
      const activeModel = candidatesToTry[i];
      const displayName = getModelDisplayName(activeModel);

      updateLoadingBubble(loadingId, `Thinking with ${displayName}...`);

      const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${activeModel}:generateContent?key=${apiKey}`;

      try {
        const response = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });

        if (response.ok) {
          const data = await response.json();
          finalReply = data.candidates?.[0]?.content?.parts?.[0]?.text;
          successfulModel = activeModel;
          break; // Success!
        }

        const errJson = await response.json().catch(() => ({}));
        const errMsg = errJson?.error?.message || `HTTP ${response.status} ${response.statusText}`;
        lastError = errMsg;

        const isDemandIssue = 
          response.status === 429 || 
          response.status === 503 || 
          response.status === 404 || 
          errMsg.toLowerCase().includes("high demand") || 
          errMsg.toLowerCase().includes("spikes in demand") || 
          errMsg.toLowerCase().includes("overloaded") || 
          errMsg.toLowerCase().includes("exhausted") ||
          errMsg.toLowerCase().includes("billing account required") ||
          errMsg.toLowerCase().includes("try again later");

        if (isDemandIssue && i < candidatesToTry.length - 1) {
          const nextModel = candidatesToTry[i + 1];
          console.warn(`Demand spike or limitation on ${activeModel}. Auto-falling back to ${nextModel}...`);
          updateLoadingBubble(loadingId, `Capacity busy on ${displayName}, auto-routing to ${getModelDisplayName(nextModel)}...`);
          await new Promise(res => setTimeout(res, 250));
          continue;
        } else {
          break;
        }
      } catch (netErr) {
        lastError = netErr.message;
        if (i < candidatesToTry.length - 1) {
          continue;
        }
        break;
      }
    }

    removeElement(loadingId);

    if (finalReply) {
      chatHistory.push({ role: "model", parts: [{ text: finalReply }] });
      
      const isFallback = selectedSetting !== "auto" && successfulModel !== selectedSetting;
      const modelTag = selectedSetting === "auto" 
        ? `⚡ ${getModelDisplayName(successfulModel)}` 
        : (isFallback ? `⚡ Fallback: ${getModelDisplayName(successfulModel)}` : getModelDisplayName(successfulModel));

      appendMessage("assistant", finalReply, null, modelTag);
    } else {
      appendMessage("assistant", `⚠️ **Gemini API Error:** ${lastError}\n\nAll available models were experiencing temporary high demand. Please try sending your message again in a moment.`);
    }

  } catch (err) {
    removeElement(loadingId);
    console.error("Gemini invocation error:", err);
    appendMessage("assistant", `⚠️ **Network / Connection Error:** ${err.message}.`);
  } finally {
    isGenerating = false;
    btnSend.disabled = promptInput.value.trim().length === 0;
  }
}

// Test Connection with Fallback Verification
async function testApiKeyConnection() {
  const key = (apiKeyInput.value || "").trim();
  if (!key || key.length < 5) {
    testStatus.innerHTML = "<span style='color:#ef4444;'>Please paste an API key first.</span>";
    return;
  }

  testStatus.innerHTML = "<span style='color:var(--text-muted);'>Testing connection to Google AI Studio...</span>";
  
  let testSuccess = false;
  let lastErr = "";
  let workingModel = "";

  for (const model of ORDERED_MODELS) {
    try {
      const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ role: "user", parts: [{ text: "ping" }] }],
          generationConfig: { maxOutputTokens: 5 }
        })
      });

      if (response.ok) {
        testSuccess = true;
        workingModel = model;
        break;
      }

      const errJson = await response.json().catch(() => ({}));
      lastErr = errJson?.error?.message || `HTTP ${response.status}`;
      if (lastErr.toLowerCase().includes("high demand") || response.status === 429 || response.status === 503) {
        continue;
      }
      break;
    } catch (err) {
      lastErr = err.message;
      break;
    }
  }

  if (testSuccess) {
    testStatus.innerHTML = `<span style='color:#10b981; font-weight:600;'>✓ Connected! (${getModelDisplayName(workingModel)} ready)</span>`;
    persistApiKey(key);
    checkApiStatus();
  } else {
    testStatus.innerHTML = `<span style='color:#ef4444;'>✗ Error: ${lastErr}</span>`;
    statusDot.className = "status-dot error";
  }
}

// UI Rendering & Transcript Persistence
function appendMessage(role, content, selectionBadgeText = null, modelBadgeText = null) {
  const timeStr = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  renderMessageNode(role, content, timeStr, selectionBadgeText, modelBadgeText);

  // Save to persistent session transcript
  savedMessages.push({
    role,
    content,
    time: timeStr,
    selection: selectionBadgeText,
    modelBadge: modelBadgeText
  });

  try {
    localStorage.setItem("gemini_chat_transcript", JSON.stringify(savedMessages.slice(-50)));
  } catch (e) {}
}

function renderMessageNode(role, content, timeStr, selectionBadgeText = null, modelBadgeText = null) {
  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${role}`;

  const bubble = document.createElement("div");
  bubble.className = "message-bubble";

  if (role === "assistant") {
    let rawHtml = marked.parse(content);
    bubble.innerHTML = rawHtml;
    enhanceCodeBlocks(bubble);
  } else {
    bubble.textContent = content;
  }

  msgDiv.appendChild(bubble);

  const meta = document.createElement("div");
  meta.className = "message-meta";
  meta.innerHTML = `<span>${timeStr || new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}</span>`;

  if (selectionBadgeText) {
    meta.innerHTML += `<span>•</span><span style="font-family:monospace;">${selectionBadgeText}</span>`;
  }
  if (modelBadgeText) {
    meta.innerHTML += `<span>•</span><span style="color:#10b981; font-weight:500;">${modelBadgeText}</span>`;
  }

  msgDiv.appendChild(meta);
  chatMessages.appendChild(msgDiv);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

function appendLoadingBubble(id) {
  const msgDiv = document.createElement("div");
  msgDiv.className = "message assistant";
  msgDiv.id = id;

  const bubble = document.createElement("div");
  bubble.className = "message-bubble";
  bubble.innerHTML = `
    <div class="loading-dots">
      <div class="dot"></div>
      <div class="dot"></div>
      <div class="dot"></div>
      <span class="loading-status" style="font-size:11px; color:var(--text-muted); margin-left:6px;">Thinking...</span>
    </div>
  `;

  msgDiv.appendChild(bubble);
  chatMessages.appendChild(msgDiv);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

function updateLoadingBubble(id, statusText) {
  const el = document.getElementById(id);
  if (!el) return;
  const span = el.querySelector(".loading-status");
  if (span) span.textContent = statusText;
}

function removeElement(id) {
  const el = document.getElementById(id);
  if (el) el.remove();
}

// Enhance code blocks with "Copy" and "Insert into Cell" buttons
function enhanceCodeBlocks(container) {
  const pres = container.querySelectorAll("pre");

  pres.forEach((pre) => {
    const code = pre.querySelector("code");
    if (!code) return;

    const rawText = code.innerText.trim();
    const isFormula = rawText.startsWith("=") || rawText.includes("XLOOKUP") || rawText.includes("SUMIFS");

    const wrapper = document.createElement("div");
    wrapper.className = "code-block-wrapper";

    const header = document.createElement("div");
    header.className = "code-header";
    header.innerHTML = `<span>${isFormula ? "Excel Formula" : "Code / Formula"}</span>`;

    const actions = document.createElement("div");
    actions.className = "code-actions";

    // Copy Button
    const copyBtn = document.createElement("button");
    copyBtn.className = "code-btn";
    copyBtn.innerHTML = "📋 Copy";
    copyBtn.addEventListener("click", () => {
      navigator.clipboard.writeText(rawText);
      copyBtn.innerHTML = "✓ Copied";
      setTimeout(() => (copyBtn.innerHTML = "📋 Copy"), 1500);
    });
    actions.appendChild(copyBtn);

    // Insert Button (if in Excel)
    if (isOfficeReady && isFormula) {
      const insertBtn = document.createElement("button");
      insertBtn.className = "code-btn btn-insert";
      insertBtn.innerHTML = "📥 Insert into Cell";
      insertBtn.addEventListener("click", async () => {
        try {
          await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.formulas = [[rawText]];
            await context.sync();
          });
          insertBtn.innerHTML = "✓ Inserted!";
          setTimeout(() => (insertBtn.innerHTML = "📥 Insert into Cell"), 1500);
        } catch (err) {
          alert(`Could not insert formula: ${err.message}`);
        }
      });
      actions.appendChild(insertBtn);
    }

    header.appendChild(actions);
    pre.parentNode.insertBefore(wrapper, pre);
    wrapper.appendChild(header);
    wrapper.appendChild(pre);
  });
}

// Download Transcript as Markdown (.md)
function downloadTranscriptMd() {
  if (!savedMessages || savedMessages.length === 0) {
    alert("No conversation messages to export yet.");
    return;
  }

  let md = `# Gemini AI for Excel — Session Transcript\n\n`;
  md += `*Exported on ${new Date().toLocaleString()}*\n`;
  md += `*Total Messages: ${savedMessages.length}*\n\n---\n\n`;

  savedMessages.forEach((msg) => {
    const speaker = msg.role === "user" ? "### 👤 User" : `### 🤖 Gemini (${msg.modelBadge || "AI"})`;
    const context = msg.selection ? `*(Range: ${msg.selection} • ${msg.time})*` : `*(${msg.time})*`;
    md += `${speaker} ${context}\n\n`;
    md += `${msg.content}\n\n---\n\n`;
  });

  const blob = new Blob([md], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Gemini_Excel_Transcript_${new Date().toISOString().slice(0, 10)}.md`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Save Transcript to Excel Worksheet ("AI_Transcript")
async function exportTranscriptToSheet() {
  if (!savedMessages || savedMessages.length === 0) {
    alert("No conversation messages to save yet.");
    return;
  }

  if (!isOfficeReady) {
    downloadTranscriptMd();
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetName = "AI_Transcript";
      let logSheet = sheets.items.find(s => s.name === sheetName);
      let startRow = 1;

      if (!logSheet) {
        logSheet = sheets.add(sheetName);
        // Style header
        const header = logSheet.getRange("A1:D1");
        header.values = [["Timestamp", "Speaker", "Excel Range", "Message / Solution"]];
        header.format.fill.color = "#107c41"; // Excel green
        header.format.font.color = "#ffffff";
        header.format.font.bold = true;
        startRow = 2;
      } else {
        const used = logSheet.getUsedRangeOrNullObject();
        await context.sync();
        if (used && !used.isNullObject) {
          used.load("rowCount");
          await context.sync();
          startRow = used.rowCount + 1;
        } else {
          startRow = 2;
        }
      }

      const rows = savedMessages.map(m => [
        m.time || new Date().toLocaleTimeString(),
        m.role === "user" ? "User" : `Gemini (${m.modelBadge || "AI"})`,
        m.selection || "-",
        m.content
      ]);

      const endRow = startRow + rows.length - 1;
      const targetRange = logSheet.getRange(`A${startRow}:D${endRow}`);
      targetRange.values = rows;
      logSheet.getRange("A:D").format.autofitColumns();
      logSheet.activate();
      await context.sync();

      alert(`✓ Successfully saved ${rows.length} messages to '${sheetName}' worksheet!`);
    });
  } catch (err) {
    console.error("Export to sheet error:", err);
    alert(`Could not write to worksheet: ${err.message}. Exporting as Markdown file instead.`);
    downloadTranscriptMd();
  }
}

function renderWelcomeCard() {
  const welcome = document.createElement("div");
  welcome.className = "welcome-card";
  welcome.innerHTML = `
    <div class="welcome-title">Gemini AI Assistant for Excel</div>
    <div class="welcome-desc">
      Direct connection to Google AI Studio. 100% Free, zero paywalls, zero middleman.
    </div>
    <div class="welcome-features">
      <div class="feature-item">⚡ <strong>Auto Smart Fallback:</strong> Automatically switches models if Google experiences demand spikes.</div>
      <div class="feature-item">💾 <strong>Session Persistence:</strong> Your chats are auto-saved across reloads and exportable to Excel or Markdown.</div>
      <div class="feature-item">📊 <strong>Dynamic Context:</strong> Detects selected cells and feeds values/formulas to Gemini.</div>
    </div>
  `;
  chatMessages.appendChild(welcome);
}
