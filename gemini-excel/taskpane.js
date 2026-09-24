/**
 * Gemini AI Assistant for Excel - Taskpane Controller
 * 100% Free, Client-side Bring-Your-Own-Key (BYOK) via Google AI Studio
 */

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

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    isOfficeReady = true;
    console.log("Running inside Microsoft Excel");
    
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
});

function getStoredApiKey() {
  return (currentApiKey || localStorage.getItem("gemini_api_key") || sessionStorage.getItem("gemini_api_key") || (apiKeyInput ? apiKeyInput.value : "") || "").trim();
}

function loadSettings() {
  currentApiKey = (localStorage.getItem("gemini_api_key") || sessionStorage.getItem("gemini_api_key") || "").trim();
  let model = localStorage.getItem("gemini_model") || "gemini-3.6-flash";
  
  // Auto-migrate deprecated 2.5 models
  if (model === "gemini-2.5-flash" || model === "gemini-2.5-pro") {
    model = "gemini-3.6-flash";
    localStorage.setItem("gemini_model", model);
  }

  const sysPrompt = localStorage.getItem("gemini_system_prompt") || DEFAULT_SYSTEM_PROMPT;
  const temp = localStorage.getItem("gemini_temp") || "0.2";

  apiKeyInput.value = currentApiKey;
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

  currentApiKey = key;
  try {
    localStorage.setItem("gemini_api_key", key);
    sessionStorage.setItem("gemini_api_key", key);
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
    localStorage.setItem("gemini_model", e.target.value);
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

  btnClearChat.addEventListener("click", () => {
    if (confirm("Clear conversation history?")) {
      chatHistory = [];
      chatMessages.innerHTML = "";
      renderWelcomeCard();
    }
  });

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

// Send Message
async function sendMessage() {
  const text = promptInput.value.trim();
  if (!text || isGenerating) return;

  const apiKey = getStoredApiKey();
  if (!apiKey || apiKey.length < 5) {
    settingsModal.classList.add("active");
    testStatus.innerHTML = "<span style='color:#ef4444;'>Please configure your Google AI Studio API Key first.</span>";
    return;
  }

  // Remove welcome card if present
  const welcomeCard = document.querySelector(".welcome-card");
  if (welcomeCard) welcomeCard.remove();

  // Add User Message to UI
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
    const selectionContext = `
[Active Selection Context in Excel:
- Address: ${currentSelection.address}
- Dimensions: ${currentSelection.rowCount} rows × ${currentSelection.columnCount} columns
- Cell Values: ${JSON.stringify(currentSelection.values)}
- Cell Formulas: ${JSON.stringify(currentSelection.formulas)}
]`;
    contextualPrompt = `${text}\n\n${selectionContext}`;
  }

  try {
    const model = modelSelector.value;
    const sysPrompt = localStorage.getItem("gemini_system_prompt") || DEFAULT_SYSTEM_PROMPT;
    const temp = parseFloat(localStorage.getItem("gemini_temp") || "0.2");

    // Build History
    chatHistory.push({ role: "user", parts: [{ text: contextualPrompt }] });

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

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

    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    removeElement(loadingId);

    if (!response.ok) {
      const errJson = await response.json().catch(() => ({}));
      const errMsg = errJson?.error?.message || `HTTP ${response.status} ${response.statusText}`;
      appendMessage("assistant", `⚠️ **Gemini API Error:** ${errMsg}\n\nPlease verify your API key or model quota in Settings.`);
      return;
    }

    const data = await response.json();
    const replyText = data.candidates?.[0]?.content?.parts?.[0]?.text || "No response received from Gemini.";

    // Store in history
    chatHistory.push({ role: "model", parts: [{ text: replyText }] });

    // Render response
    appendMessage("assistant", replyText);

  } catch (err) {
    removeElement(loadingId);
    console.error("Gemini invocation error:", err);
    appendMessage("assistant", `⚠️ **Network / Connection Error:** ${err.message}. Check your internet connection.`);
  } finally {
    isGenerating = false;
    btnSend.disabled = promptInput.value.trim().length === 0;
  }
}

// Test Connection
async function testApiKeyConnection() {
  const key = apiKeyInput.value.trim();
  if (!key) {
    testStatus.innerHTML = "<span style='color:#ef4444;'>Please paste an API key first.</span>";
    return;
  }

  testStatus.innerHTML = "<span style='color:var(--text-muted);'>Testing connection to Google AI Studio...</span>";
  const model = modelSelector.value;

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
      testStatus.innerHTML = "<span style='color:#10b981; font-weight:600;'>✓ Connected successfully! Model is ready.</span>";
      statusDot.className = "status-dot connected";
    } else {
      const err = await response.json().catch(() => ({}));
      testStatus.innerHTML = `<span style='color:#ef4444;'>✗ Error (${response.status}): ${err?.error?.message || "Invalid Key"}</span>`;
      statusDot.className = "status-dot error";
    }
  } catch (err) {
    testStatus.innerHTML = `<span style='color:#ef4444;'>✗ Network failed: ${err.message}</span>`;
  }
}

// UI Rendering
function appendMessage(role, content, selectionBadgeText = null) {
  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${role}`;

  const bubble = document.createElement("div");
  bubble.className = "message-bubble";

  if (role === "assistant") {
    // Parse Markdown using marked
    let rawHtml = marked.parse(content);
    bubble.innerHTML = rawHtml;
    enhanceCodeBlocks(bubble);
  } else {
    bubble.textContent = content;
  }

  msgDiv.appendChild(bubble);

  const meta = document.createElement("div");
  meta.className = "message-meta";
  const time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  meta.innerHTML = `<span>${time}</span>`;

  if (selectionBadgeText) {
    meta.innerHTML += `<span>•</span><span style="font-family:monospace;">${selectionBadgeText}</span>`;
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
      <span style="font-size:11px; color:var(--text-muted); margin-left:6px;">Thinking...</span>
    </div>
  `;

  msgDiv.appendChild(bubble);
  chatMessages.appendChild(msgDiv);
  chatMessages.scrollTop = chatMessages.scrollHeight;
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

function renderWelcomeCard() {
  const welcome = document.createElement("div");
  welcome.className = "welcome-card";
  welcome.innerHTML = `
    <div class="welcome-title">Gemini AI Assistant for Excel</div>
    <div class="welcome-desc">
      Direct connection to Google AI Studio. 100% Free, zero paywalls, zero middleman.
    </div>
    <div class="welcome-features">
      <div class="feature-item">⚡ <strong>One-Click Formulas:</strong> Insert XLOOKUP, LET & dynamic arrays.</div>
      <div class="feature-item">📊 <strong>Selection Aware:</strong> Highlights & analyzes active cells automatically.</div>
      <div class="feature-item">🔒 <strong>100% Private BYOK:</strong> Key stays in local storage on your device.</div>
    </div>
  `;
  chatMessages.appendChild(welcome);
}
