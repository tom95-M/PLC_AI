const toolGrid = document.getElementById("toolGrid");
const template = document.getElementById("toolCardTemplate");
const refreshBtn = document.getElementById("refreshBtn");
const aiInput = document.getElementById("aiInput");
const aiSendBtn = document.getElementById("aiSendBtn");
const chatHistory = document.getElementById("chatHistory");
const chatClearBtn = document.getElementById("chatClearBtn");

const openAgentModalBtn = document.getElementById("openAgentModalBtn");
const closeAgentModalBtn = document.getElementById("closeAgentModalBtn");
const agentModal = document.getElementById("agentModal");

const agentSelect = document.getElementById("agentSelect");
const agentDesc = document.getElementById("agentDesc");
const agentNewBtn = document.getElementById("agentNewBtn");
const agentSaveBtn = document.getElementById("agentSaveBtn");
const agentDeleteBtn = document.getElementById("agentDeleteBtn");
const agentBaseUrl = document.getElementById("agentBaseUrl");
const agentModel = document.getElementById("agentModel");
const agentTemp = document.getElementById("agentTemp");
const agentApiKey = document.getElementById("agentApiKey");
const agentLlmEnabled = document.getElementById("agentLlmEnabled");
const agentTestBtn = document.getElementById("agentTestBtn");

const state = {
  tools: [],
  cardRefs: new Map(),
  agents: [],
  currentAgentId: "",
  messages: [],
};

function appendLog(text, isError = false) {
  addChat("system", text, isError);
}

function openAgentModal() {
  if (!agentModal) return;
  agentModal.classList.remove("hidden");
  agentModal.setAttribute("aria-hidden", "false");
}

function closeAgentModal() {
  if (!agentModal) return;
  agentModal.classList.add("hidden");
  agentModal.setAttribute("aria-hidden", "true");
}

function addChat(role, content, isError = false) {
  state.messages.push({ role, content, isError: !!isError });
  if (state.messages.length > 80) state.messages = state.messages.slice(-80);
  renderChat();
}

function renderChat() {
  chatHistory.innerHTML = "";
  for (const msg of state.messages) {
    const item = document.createElement("div");
    item.className = "chat-item";
    if (msg.role === "user") {
      item.classList.add("chat-user");
    } else if (msg.role === "assistant") {
      item.classList.add("chat-assistant");
    } else {
      item.classList.add("chat-system");
    }
    if (msg.isError) item.classList.add("chat-error");
    item.textContent = msg.content;
    chatHistory.appendChild(item);
  }
  chatHistory.scrollTop = chatHistory.scrollHeight;
}

async function request(url, options = {}) {
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  let data = {};
  try {
    data = await response.json();
  } catch {
    data = {};
  }
  if (!response.ok) throw new Error(data.message || "请求失败");
  return data;
}

function toolStatusClass(tool) {
  if (tool.running) return "status-running";
  if (tool.pathExists) return "status-installed";
  return "status-missing";
}

function toolStatusText(tool) {
  if (tool.running) return "运行中";
  if (tool.pathExists) return "已安装";
  return "未安装";
}

function updateToolCard(tool) {
  const refs = state.cardRefs.get(tool.id);
  if (!refs) return;
  refs.card.className = `tool-chip ${toolStatusClass(tool)}`;
  refs.card.title = `${tool.name} · ${toolStatusText(tool)}`;
}

function renderToolCards(tools) {
  toolGrid.innerHTML = "";
  state.cardRefs.clear();
  for (const tool of tools) {
    const node = template.content.cloneNode(true);
    const card = node.querySelector(".tool-chip");
    const name = node.querySelector(".tool-name");
    name.textContent = tool.name || tool.id;
    toolGrid.appendChild(node);
    state.cardRefs.set(tool.id, { card, name });
    updateToolCard(tool);
  }
}

async function refreshTools() {
  const data = await request("/api/tools");
  state.tools = data.tools || [];
  if (!state.cardRefs.size) {
    renderToolCards(state.tools);
  } else {
    for (const tool of state.tools) updateToolCard(tool);
  }
}

function findAgent(agentId) {
  return state.agents.find((item) => item.id === agentId);
}

function fillAgentForm(agent) {
  if (!agent) return;
  agentDesc.value = agent.description || "";
  agentBaseUrl.value = agent.llmBaseUrl || "https://api.openai.com/v1";
  agentModel.value = agent.llmModel || "gpt-5-codex";
  agentTemp.value = String(agent.llmTemperature ?? 0.2);
  agentApiKey.value = agent.llmApiKey || "";
  agentLlmEnabled.checked = !!agent.llmEnabled;
}

function renderAgentSelect() {
  agentSelect.innerHTML = state.agents.map((agent) => `<option value="${agent.id}">${agent.name}</option>`).join("");
  agentSelect.value = state.currentAgentId;
  fillAgentForm(findAgent(state.currentAgentId));
}

async function refreshAgents() {
  const data = await request("/api/agents");
  state.agents = data.agents || [];
  state.currentAgentId = data.currentAgentId || (state.agents[0] ? state.agents[0].id : "");
  renderAgentSelect();
}

function buildAgentPayload(existingId = "") {
  const base = findAgent(existingId) || {};
  const id = existingId || `agent_${Date.now()}`;
  const name = base.name || `新Agent_${new Date().toISOString().slice(11, 19).replaceAll(":", "")}`;
  return {
    id,
    name,
    description: agentDesc.value.trim(),
    defaultTool: base.defaultTool || "",
    mode: base.mode || "balanced",
    autoCreateOnWrite: !!base.autoCreateOnWrite,
    projectPrefix: base.projectPrefix || "PLC_Project",
    llmEnabled: agentLlmEnabled.checked,
    llmProvider: "openai",
    llmBaseUrl: agentBaseUrl.value.trim() || "https://api.openai.com/v1",
    llmApiKey: agentApiKey.value.trim(),
    llmModel: agentModel.value.trim() || "gpt-5-codex",
    llmTemperature: Number(agentTemp.value || 0.2),
  };
}

async function selectAgent(agentId) {
  await request("/api/agents/select", { method: "POST", body: JSON.stringify({ agentId }) });
  state.currentAgentId = agentId;
  fillAgentForm(findAgent(agentId));
  appendLog(`已切换 Agent: ${findAgent(agentId)?.name || agentId}`);
}

async function saveAgent(isNew) {
  const payload = buildAgentPayload(isNew ? "" : state.currentAgentId);
  const data = await request("/api/agents/upsert", { method: "POST", body: JSON.stringify({ agent: payload }) });
  await refreshAgents();
  state.currentAgentId = data.currentAgentId;
  agentSelect.value = state.currentAgentId;
  appendLog(`Agent 已保存: ${data.agent.name}`);
}

async function deleteAgent() {
  await request("/api/agents/delete", { method: "POST", body: JSON.stringify({ agentId: state.currentAgentId }) });
  await refreshAgents();
  appendLog("Agent 已删除");
}

async function testLlm() {
  const payload = buildAgentPayload(state.currentAgentId);
  const data = await request("/api/llm/test", { method: "POST", body: JSON.stringify({ agent: payload }) });
  appendLog(`API 测试成功: ${data.message}`);
}

async function sendAiCommand() {
  const prompt = aiInput.value.trim();
  if (!prompt) {
    appendLog("请先输入指令", true);
    return;
  }
  addChat("user", prompt);
  aiInput.value = "";
  aiSendBtn.disabled = true;
  try {
    const data = await request("/api/ai-command", {
      method: "POST",
      body: JSON.stringify({
        prompt,
        agentId: state.currentAgentId,
        history: state.messages.filter((x) => x.role === "user" || x.role === "assistant"),
      }),
    });
    const reply = data.message || "执行完成";
    addChat("assistant", reply);
    await refreshTools();
  } catch (error) {
    addChat("assistant", `错误: ${error.message}`, true);
    await refreshTools();
  } finally {
    aiSendBtn.disabled = false;
  }
}

refreshBtn.addEventListener("click", async () => {
  try {
    await refreshTools();
    appendLog("状态已刷新");
  } catch (error) {
    appendLog(error.message, true);
  }
});

chatClearBtn.addEventListener("click", () => {
  state.messages = [];
  renderChat();
  appendLog("会话已清空");
});

if (openAgentModalBtn) openAgentModalBtn.addEventListener("click", openAgentModal);
if (closeAgentModalBtn) closeAgentModalBtn.addEventListener("click", closeAgentModal);
if (agentModal) {
  agentModal.addEventListener("click", (event) => {
    if (event.target === agentModal) closeAgentModal();
  });
}

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && agentModal && !agentModal.classList.contains("hidden")) closeAgentModal();
});

agentSelect.addEventListener("change", async (event) => {
  try {
    await selectAgent(event.target.value);
  } catch (error) {
    appendLog(error.message, true);
  }
});

agentNewBtn.addEventListener("click", async () => {
  try {
    await saveAgent(true);
  } catch (error) {
    appendLog(error.message, true);
  }
});

agentSaveBtn.addEventListener("click", async () => {
  try {
    await saveAgent(false);
  } catch (error) {
    appendLog(error.message, true);
  }
});

agentDeleteBtn.addEventListener("click", async () => {
  try {
    await deleteAgent();
  } catch (error) {
    appendLog(error.message, true);
  }
});

agentTestBtn.addEventListener("click", async () => {
  try {
    await testLlm();
  } catch (error) {
    appendLog(`API 测试失败: ${error.message}`, true);
  }
});

aiSendBtn.addEventListener("click", sendAiCommand);
aiInput.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") sendAiCommand();
});

(async () => {
  try {
    await refreshTools();
    await refreshAgents();
    renderChat();
    appendLog("系统已就绪");
  } catch (error) {
    appendLog(error.message, true);
  }
})();
