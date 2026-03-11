import csv
import json
import os
import re
import subprocess
import time
import urllib.request
import urllib.error
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from io import StringIO
from pathlib import Path
from urllib.parse import urlparse
import winreg


ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
HOST = "127.0.0.1"
PORT = 9527
AGENTS_FILE = ROOT / "agents.json"

TOOLS = {
    "autoshop": {
        "name": "AutoShop",
        "vendor": "Inovance",
        "process_names": ["AutoShop.exe"],
        "exe_names": ["AutoShop.exe"],
        "candidates": [
            r"C:\Program Files\Inovance\AutoShop\AutoShop.exe",
            r"C:\Program Files (x86)\Inovance\AutoShop\AutoShop.exe",
            r"C:\Program Files\Inovance\AutoShop\Bin\AutoShop.exe",
            r"C:\Program Files (x86)\Inovance\AutoShop\Bin\AutoShop.exe",
            r"D:\Inovance Control\AutoShop\AutoShop.exe",
        ],
        "registry_keywords": ["autoshop", "inovance"],
    },
    "gxwork2": {
        "name": "GX Works2",
        "vendor": "Mitsubishi",
        "process_names": ["GXWorks2.exe", "GD2.exe", "GX Works2 Startup.exe"],
        "exe_names": ["GXWorks2.exe", "GD2.exe", "GX Works2 Startup.exe"],
        "candidates": [
            r"C:\Program Files (x86)\MELSOFT\GX Works2\GXWorks2.exe",
            r"C:\Program Files\MELSOFT\GX Works2\GXWorks2.exe",
            r"C:\Program Files (x86)\MELSOFT\GPPW2\GD2.exe",
            r"C:\Program Files (x86)\MELSOFT\GPPW2\GX Works2 Startup.exe",
            r"D:\Program Files (x86)\MELSOFT\GX Works2\GXWorks2.exe",
            r"D:\Program Files\MELSOFT\GX Works2\GXWorks2.exe",
            r"D:\Program Files (x86)\MELSOFT\GPPW2\GD2.exe",
            r"D:\Program Files (x86)\MELSOFT\GPPW2\GX Works2 Startup.exe",
        ],
        "registry_keywords": ["gx works2", "melsoft", "mitsubishi"],
    },
    "gxwork3": {
        "name": "GX Works3",
        "vendor": "Mitsubishi",
        "process_names": ["GXW3.exe"],
        "exe_names": ["GXW3.exe"],
        "candidates": [
            r"C:\Program Files\MELSOFT\GX Works3\GXW3.exe",
            r"C:\Program Files (x86)\MELSOFT\GX Works3\GXW3.exe",
            r"D:\GXWORK3\GPPW3\GXW3.exe",
        ],
        "registry_keywords": ["gx works3", "melsoft", "mitsubishi"],
    },
    "tiaportalv18": {
        "name": "TIA PortalV18",
        "vendor": "Siemens",
        "process_names": ["Siemens.Automation.Portal.exe", "Portal.exe"],
        "exe_names": ["Siemens.Automation.Portal.exe", "Portal.exe"],
        "candidates": [
            r"C:\Program Files\Siemens\Automation\Portal V18\Bin\Siemens.Automation.Portal.exe",
            r"C:\Program Files\Siemens\Automation\Portal V18\Bin\Portal.exe",
            r"D:\Program Files\Siemens\Automation\Portal V18\Bin\Siemens.Automation.Portal.exe",
            r"D:\Program Files\Siemens\Automation\Portal V18\Bin\Portal.exe",
        ],
        "registry_keywords": ["portal v18", "tia portal", "siemens"],
    },
    "inoproshop": {
        "name": "InoProShop",
        "vendor": "Inovance",
        "process_names": ["InoProShop.exe", "InoproShop.exe"],
        "exe_names": ["InoProShop.exe", "InoproShop.exe"],
        "candidates": [
            r"C:\Program Files\Inovance\InoProShop\InoProShop.exe",
            r"C:\Program Files (x86)\Inovance\InoProShop\InoProShop.exe",
        ],
        "registry_keywords": ["inoproshop", "inovance"],
    },
}

REGISTRY_UNINSTALL_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
]

AI_TOOL_ALIASES = {
    "autoshop": ["autoshop", "auto shop", "huichuan", "汇川"],
    "gxwork2": ["gxwork2", "gx works2", "gx2", "gppw2", "gd2"],
    "gxwork3": ["gxwork3", "gx works3", "gx3", "gppw3", "gxw3"],
    "tiaportalv18": ["tiaportal", "tia portal", "portal v18", "tia v18", "siemens"],
    "inoproshop": ["inoproshop", "ino pro shop"],
}

WINDOW_TITLE_HINTS = {
    "autoshop": "AutoShop",
    "gxwork2": "GX Works2",
    "gxwork3": "GX Works3",
    "tiaportalv18": "TIA Portal",
    "inoproshop": "InoProShop",
}

PROCESS_NAME_HINTS = {
    "autoshop": ["AutoShop"],
    "gxwork2": ["GXWorks2", "GD2"],
    "gxwork3": ["GXW3"],
    "tiaportalv18": ["Siemens.Automation.Portal", "Portal"],
    "inoproshop": ["InoProShop", "InoproShop"],
}

DEFAULT_AGENTS = [
    {
        "id": "default",
        "name": "通用Agent",
        "description": "Rule-based + optional LLM API",
        "defaultTool": "",
        "mode": "balanced",
        "autoCreateOnWrite": False,
        "projectPrefix": "PLC_Project",
        "llmEnabled": False,
        "llmProvider": "openai",
        "llmBaseUrl": "https://api.openai.com/v1",
        "llmApiKey": "",
        "llmModel": "gpt-4.1-mini",
        "llmTemperature": 0.2,
    }
]

LAST_CODE_DRAFT = ""
AGENT_STATE = {"agents": [], "currentAgentId": "default"}


def run_hidden(command):
    kwargs = {"capture_output": True, "text": True, "check": False, "encoding": "utf-8", "errors": "ignore"}
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kwargs["startupinfo"] = startupinfo
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.run(command, **kwargs)


def normalize_agent(data):
    raw_id = str(data.get("id", "")).strip().lower() or ("agent_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    clean_id = re.sub(r"[^a-z0-9_\-]", "_", raw_id)
    default_tool = str(data.get("defaultTool", "")).strip().lower()
    if default_tool and default_tool not in TOOLS:
        default_tool = ""
    mode = str(data.get("mode", "balanced")).strip().lower()
    if mode not in ("safe", "balanced", "aggressive"):
        mode = "balanced"
    provider = str(data.get("llmProvider", "openai")).strip().lower()
    if provider not in ("openai",):
        provider = "openai"
    base_url = str(data.get("llmBaseUrl", "https://api.openai.com/v1")).strip().rstrip("/")
    model = str(data.get("llmModel", "gpt-4.1-mini")).strip() or "gpt-4.1-mini"
    try:
        temperature = float(data.get("llmTemperature", 0.2))
    except Exception:
        temperature = 0.2
    temperature = max(0.0, min(2.0, temperature))
    return {
        "id": clean_id,
        "name": str(data.get("name", "未命名Agent")).strip() or "未命名Agent",
        "description": str(data.get("description", "")).strip(),
        "defaultTool": default_tool,
        "mode": mode,
        "autoCreateOnWrite": bool(data.get("autoCreateOnWrite", False)),
        "projectPrefix": str(data.get("projectPrefix", "PLC_Project")).strip() or "PLC_Project",
        "llmEnabled": bool(data.get("llmEnabled", False)),
        "llmProvider": provider,
        "llmBaseUrl": base_url,
        "llmApiKey": str(data.get("llmApiKey", "")).strip(),
        "llmModel": model,
        "llmTemperature": temperature,
    }


def load_agents():
    if not AGENTS_FILE.exists():
        AGENT_STATE["agents"] = list(DEFAULT_AGENTS)
        AGENT_STATE["currentAgentId"] = "default"
        save_agents()
        return
    try:
        raw = json.loads(AGENTS_FILE.read_text(encoding="utf-8"))
        agents = raw.get("agents", [])
        if not agents:
            raise ValueError("empty agents")
        AGENT_STATE["agents"] = [normalize_agent(item) for item in agents]
        AGENT_STATE["currentAgentId"] = raw.get("currentAgentId", AGENT_STATE["agents"][0]["id"])
    except Exception:
        AGENT_STATE["agents"] = list(DEFAULT_AGENTS)
        AGENT_STATE["currentAgentId"] = "default"
        save_agents()


def save_agents():
    payload = {"agents": AGENT_STATE["agents"], "currentAgentId": AGENT_STATE["currentAgentId"]}
    AGENTS_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def get_agent(agent_id):
    for item in AGENT_STATE["agents"]:
        if item.get("id") == agent_id:
            return item
    return None


def get_current_agent():
    current = get_agent(AGENT_STATE["currentAgentId"])
    if current:
        return current
    fallback = AGENT_STATE["agents"][0]
    AGENT_STATE["currentAgentId"] = fallback["id"]
    return fallback


def path_exists(path):
    return bool(path and Path(path).exists())


def parse_registry_path(value):
    return str(value).strip().strip('"').split(",")[0].strip().strip('"')


def process_running(process_names):
    try:
        result = run_hidden(["tasklist", "/FO", "CSV", "/NH"])
    except Exception:
        return False
    known = {name.lower() for name in process_names}
    reader = csv.reader(StringIO(result.stdout))
    for row in reader:
        if row and row[0].strip().lower() in known:
            return True
    return False


def read_registry_candidates(keywords, exe_names):
    candidates = []
    lowered_keywords = [k.lower() for k in keywords]
    lowered_exe_names = [e.lower() for e in exe_names]
    for hive, uninstall_path in REGISTRY_UNINSTALL_PATHS:
        try:
            root_key = winreg.OpenKey(hive, uninstall_path)
        except OSError:
            continue
        index = 0
        while True:
            try:
                sub_key_name = winreg.EnumKey(root_key, index)
            except OSError:
                break
            index += 1
            try:
                sub_key = winreg.OpenKey(root_key, sub_key_name)
            except OSError:
                continue
            try:
                display_name = str(winreg.QueryValueEx(sub_key, "DisplayName")[0]).lower()
            except OSError:
                continue
            if not any(keyword in display_name for keyword in lowered_keywords):
                continue
            for value_name in ["DisplayIcon", "InstallLocation"]:
                try:
                    raw = winreg.QueryValueEx(sub_key, value_name)[0]
                except OSError:
                    continue
                raw_path = Path(parse_registry_path(raw))
                if raw_path.is_file():
                    if raw_path.suffix.lower() == ".exe" and (not lowered_exe_names or raw_path.name.lower() in lowered_exe_names):
                        candidates.append(str(raw_path))
                    else:
                        parent = raw_path.parent
                        for exe_name in exe_names:
                            candidates.append(str(parent / exe_name))
                            candidates.append(str(parent / "Bin" / exe_name))
                elif raw_path.is_dir():
                    for exe_name in exe_names:
                        candidates.append(str(raw_path / exe_name))
                        candidates.append(str(raw_path / "Bin" / exe_name))
    return candidates


def discover_tool_path(tool_id):
    tool = TOOLS[tool_id]
    candidates = list(tool["candidates"])
    candidates.extend(read_registry_candidates(tool["registry_keywords"], tool["exe_names"]))
    seen = set()
    for candidate in candidates:
        normalized = os.path.normpath(candidate)
        if normalized in seen:
            continue
        seen.add(normalized)
        if path_exists(normalized):
            return normalized
    return ""


def build_tool_status(tool_id):
    tool = TOOLS[tool_id]
    detected = discover_tool_path(tool_id)
    return {"id": tool_id, "name": tool["name"], "vendor": tool["vendor"], "detectedPath": detected, "pathExists": bool(detected), "running": process_running(tool["process_names"])}


def launch_tool(tool_id):
    detected = discover_tool_path(tool_id)
    if not detected:
        return False, "Cannot auto-detect install path."
    try:
        os.startfile(detected)  # type: ignore[attr-defined]
        return True, "Launch triggered."
    except OSError as exc:
        return False, f"Launch failed: {exc}"


def _ps_escape_single_quoted(value):
    return value.replace("'", "''")


def _escape_sendkeys_text(value):
    mapping = {"{": "{{}", "}": "{}}", "+": "{+}", "^": "{^}", "%": "{%}", "~": "{~}", "(": "{(}", ")": "{)}", "[": "{[}", "]": "{]}"}
    return "".join(mapping.get(ch, ch) for ch in value)


def _run_powershell(script):
    return run_hidden(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script])


def activate_window(tool_id, retries=12, delay=0.5):
    hint = WINDOW_TITLE_HINTS.get(tool_id, TOOLS[tool_id]["name"])
    escaped = _ps_escape_single_quoted(hint)
    proc_patterns = PROCESS_NAME_HINTS.get(tool_id, [])
    proc_match_expr = " -or ".join(f"$_.ProcessName -like '*{_ps_escape_single_quoted(p)}*'" for p in proc_patterns) or "$false"
    for _ in range(retries):
        script = (
            "$ws=New-Object -ComObject WScript.Shell;"
            f"$p=Get-Process | Where-Object {{ ($_.MainWindowHandle -ne 0) -and ({proc_match_expr}) }} | Sort-Object StartTime -Descending | Select-Object -First 1;"
            "if ($p -and $ws.AppActivate([int]$p.Id)) { exit 0 };"
            f"if ($ws.AppActivate('{escaped}')) {{ exit 0 }} else {{ exit 2 }}"
        )
        if _run_powershell(script).returncode == 0:
            return True
        time.sleep(delay)
    return False


def sanitize_project_name(raw):
    cleaned = re.sub(r"[^\w\-]", "", raw, flags=re.UNICODE).strip("_-")
    return cleaned or ("PLC_Project_" + datetime.now().strftime("%Y%m%d_%H%M%S"))


def extract_project_name(prompt, prefix):
    patterns = [r"项目名[:：\s]*([A-Za-z0-9_\-\u4e00-\u9fff]+)", r"项目[:：\s]*([A-Za-z0-9_\-\u4e00-\u9fff]+)", r"project\s*[:：]?\s*([A-Za-z0-9_\-]+)"]
    for pattern in patterns:
        match = re.search(pattern, prompt, flags=re.IGNORECASE)
        if match:
            return sanitize_project_name(match.group(1))
    return sanitize_project_name(f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")


def create_project_in_editor(tool_id, project_name):
    ok, msg = launch_tool(tool_id)
    if not ok:
        return False, msg
    if not activate_window(tool_id):
        return False, "Tool opened but window activation failed."
    safe_name = _ps_escape_single_quoted(project_name)
    script = "$ws=New-Object -ComObject WScript.Shell;Start-Sleep -Milliseconds 300;$ws.SendKeys('^n');Start-Sleep -Milliseconds 700;" + f"$ws.SendKeys('{safe_name}');" + "Start-Sleep -Milliseconds 100;$ws.SendKeys('{ENTER}');"
    if _run_powershell(script).returncode != 0:
        return False, "Create project automation failed."
    return True, f"Project create triggered: {project_name}"


def type_code_to_editor(tool_id, code, replace_existing):
    if not activate_window(tool_id):
        return False, "Editor window not active."
    lines = code.splitlines()
    chunks = ["$ws=New-Object -ComObject WScript.Shell;", "Start-Sleep -Milliseconds 220;"]
    if replace_existing:
        chunks += ["$ws.SendKeys('^a');", "Start-Sleep -Milliseconds 80;", "$ws.SendKeys('{DEL}');", "Start-Sleep -Milliseconds 90;"]
    for idx, line in enumerate(lines):
        escaped = _ps_escape_single_quoted(_escape_sendkeys_text(line))
        chunks.append(f"$ws.SendKeys('{escaped}');")
        if idx < len(lines) - 1:
            chunks.append("$ws.SendKeys('{ENTER}');")
        chunks.append("Start-Sleep -Milliseconds 35;")
    if _run_powershell("".join(chunks)).returncode != 0:
        return False, "Type-in automation failed."
    return True, "Code typed into editor."


def normalize_for_match(text):
    return re.sub(r"[\s_\-]+", "", text.strip().lower())


def resolve_tools_from_prompt(prompt):
    normalized = normalize_for_match(prompt)
    matched = []
    for tool_id, tool in TOOLS.items():
        tokens = [tool_id, tool["name"]]
        tokens.extend(AI_TOOL_ALIASES.get(tool_id, []))
        if any(normalize_for_match(token) in normalized for token in tokens):
            matched.append(tool_id)
    return matched


def pick_tool_from_context(prompt, agent):
    targets = resolve_tools_from_prompt(prompt)
    if targets:
        return targets[0]
    default_tool = str(agent.get("defaultTool", "")).strip().lower()
    if default_tool in TOOLS:
        return default_tool
    running = [tid for tid, meta in TOOLS.items() if process_running(meta["process_names"])]
    if running:
        return running[0]
    installed = [tid for tid in TOOLS if build_tool_status(tid)["pathExists"]]
    return installed[0] if installed else ""


def call_openai_responses(agent, prompt, system_instruction, history=None):
    base_url = str(agent.get("llmBaseUrl", "https://api.openai.com/v1")).rstrip("/")
    api_key = str(agent.get("llmApiKey", "")).strip()
    if not api_key:
        raise ValueError("API key is empty")
    model = str(agent.get("llmModel", "gpt-4.1-mini"))
    temperature = float(agent.get("llmTemperature", 0.2))
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}

    # SiliconFlow and many OpenAI-compatible gateways prefer chat/completions.
    prefer_chat = "siliconflow.cn" in base_url.lower()
    endpoints = ["chat/completions", "responses"] if prefer_chat else ["responses", "chat/completions"]

    last_error = None
    recent_history = []
    if isinstance(history, list):
        for item in history[-12:]:
            role = str(item.get("role", "")).strip().lower()
            content = str(item.get("content", "")).strip()
            if role in ("user", "assistant") and content:
                recent_history.append({"role": role, "content": content})

    for endpoint in endpoints:
        try:
            if endpoint == "responses":
                inputs = [{"role": "system", "content": system_instruction}]
                inputs.extend(recent_history)
                inputs.append({"role": "user", "content": prompt})
                payload = {
                    "model": model,
                    "input": inputs,
                    "temperature": temperature,
                }
                req = urllib.request.Request(
                    f"{base_url}/responses",
                    data=json.dumps(payload).encode("utf-8"),
                    headers=headers,
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=45) as resp:
                    data = json.loads(resp.read().decode("utf-8", errors="ignore"))
                if data.get("output_text"):
                    return str(data["output_text"])
                output = data.get("output", [])
                chunks = []
                for item in output:
                    for content in item.get("content", []):
                        if content.get("text"):
                            chunks.append(content["text"])
                text = "\n".join(chunks).strip()
                if text:
                    return text
            else:
                messages = [{"role": "system", "content": system_instruction}]
                messages.extend(recent_history)
                messages.append({"role": "user", "content": prompt})
                payload = {
                    "model": model,
                    "messages": messages,
                    "temperature": temperature,
                }
                req = urllib.request.Request(
                    f"{base_url}/chat/completions",
                    data=json.dumps(payload).encode("utf-8"),
                    headers=headers,
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=45) as resp:
                    data = json.loads(resp.read().decode("utf-8", errors="ignore"))
                choices = data.get("choices", [])
                if choices:
                    content = choices[0].get("message", {}).get("content", "")
                    if isinstance(content, str) and content.strip():
                        return content.strip()
        except Exception as exc:
            last_error = exc
            continue
    if last_error:
        raise last_error
    raise RuntimeError("LLM response is empty")


def generate_code(prompt, agent, history=None):
    if not bool(agent.get("llmEnabled", False)):
        raise RuntimeError("LLM API is disabled in current agent.")
    if not str(agent.get("llmApiKey", "")).strip():
        raise RuntimeError("LLM API key is empty.")
    return call_openai_responses(agent, prompt, "You generate PLC ST code. Return only code.", history=history)


def ai_handle_command(prompt, agent, history=None):
    global LAST_CODE_DRAFT
    text = str(prompt or "").strip()
    if not text:
        return {"ok": False, "message": "Please enter a command.", "actions": []}
    normalized = normalize_for_match(text)
    status_keywords = ["status", "scan", "state", "health", "状态", "扫描", "检测"]
    launch_keywords = ["open", "launch", "start", "run", "打开", "启动", "运行"]
    code_keywords = ["code", "program", "write", "st", "ladder", "代码", "程序", "编写", "写程序", "写入程序"]
    create_keywords = ["createproject", "newproject", "project", "创建项目", "新建项目", "工程"]
    paste_keywords = ["paste", "粘贴程序", "再次粘贴", "补写程序"]
    all_keywords = ["all", "全部", "所有"]
    has_create = any(item in normalized for item in create_keywords)
    has_code = any(item in normalized for item in code_keywords)
    has_all = any(item in normalized for item in all_keywords)
    mode = agent.get("mode", "balanced")

    if any(item in normalized for item in status_keywords):
        tools = [build_tool_status(tid) for tid in TOOLS]
        lines = [f"- {t['name']}: {'installed' if t['pathExists'] else 'missing'}, {'running' if t['running'] else 'stopped'}" for t in tools]
        return {"ok": True, "message": "Tool status:\n" + "\n".join(lines), "actions": [], "tools": tools}

    if any(item in normalized for item in paste_keywords):
        tool_id = pick_tool_from_context(text, agent)
        if not tool_id:
            return {"ok": False, "message": "No target tool found.", "actions": []}
        if not LAST_CODE_DRAFT:
            return {"ok": False, "message": "No draft code. Run write command first.", "actions": []}
        ok, msg = type_code_to_editor(tool_id, LAST_CODE_DRAFT, replace_existing=False)
        return {"ok": ok, "message": f"{TOOLS[tool_id]['name']}: {msg}", "actions": [{"tool": tool_id, "ok": ok, "message": msg}]}

    if has_create and has_code:
        tool_id = pick_tool_from_context(text, agent)
        if not tool_id:
            return {"ok": False, "message": "No target tool found.", "actions": []}
        project_name = extract_project_name(text, agent.get("projectPrefix", "PLC_Project"))
        ok_create, msg_create = create_project_in_editor(tool_id, project_name)
        if not ok_create:
            return {"ok": False, "message": f"{TOOLS[tool_id]['name']}: {msg_create}", "actions": [{"tool": tool_id, "ok": False, "message": msg_create}]}
        time.sleep(1.0)
        try:
            code = generate_code(text, agent, history=history)
        except Exception as exc:
            return {"ok": False, "message": f"LLM code generation failed: {type(exc).__name__}: {exc}", "actions": []}
        LAST_CODE_DRAFT = code
        ok_code, msg_code = type_code_to_editor(tool_id, code, replace_existing=(mode != "safe"))
        return {"ok": ok_code, "message": f"{TOOLS[tool_id]['name']}: {msg_create}; {msg_code}", "actions": [{"tool": tool_id, "ok": ok_create, "message": msg_create}, {"tool": tool_id, "ok": ok_code, "message": msg_code}]}

    if has_code:
        tool_id = pick_tool_from_context(text, agent)
        if not tool_id:
            return {"ok": False, "message": "Please specify target tool first.", "actions": []}
        if agent.get("autoCreateOnWrite", False):
            project_name = extract_project_name(text, agent.get("projectPrefix", "PLC_Project"))
            create_project_in_editor(tool_id, project_name)
            time.sleep(0.8)
        try:
            code = generate_code(text, agent, history=history)
        except Exception as exc:
            return {"ok": False, "message": f"LLM code generation failed: {type(exc).__name__}: {exc}", "actions": []}
        LAST_CODE_DRAFT = code
        ok, msg = type_code_to_editor(tool_id, code, replace_existing=(mode != "safe"))
        return {"ok": ok, "message": f"{TOOLS[tool_id]['name']}: {msg}", "actions": [{"tool": tool_id, "ok": ok, "message": msg}]}

    if has_create:
        tool_id = pick_tool_from_context(text, agent)
        if not tool_id:
            return {"ok": False, "message": "Please specify target tool.", "actions": []}
        project_name = extract_project_name(text, agent.get("projectPrefix", "PLC_Project"))
        ok, msg = create_project_in_editor(tool_id, project_name)
        return {"ok": ok, "message": f"{TOOLS[tool_id]['name']}: {msg}", "actions": [{"tool": tool_id, "ok": ok, "message": msg}]}

    if any(item in normalized for item in launch_keywords):
        if has_all:
            targets = [item["id"] for item in [build_tool_status(tid) for tid in TOOLS] if item["pathExists"]]
        else:
            selected = pick_tool_from_context(text, agent)
            targets = [selected] if selected else []
        if not targets:
            return {"ok": False, "message": "No launch target found.", "actions": []}
        actions = []
        for tool_id in targets:
            ok, msg = launch_tool(tool_id)
            actions.append({"tool": tool_id, "ok": ok, "message": msg})
        return {"ok": all(item["ok"] for item in actions), "message": "Launch completed." if all(item["ok"] for item in actions) else "Some launch actions failed.", "actions": actions}

    if bool(agent.get("llmEnabled", False)) and agent.get("llmApiKey"):
        try:
            reply = call_openai_responses(
                agent,
                text,
                "You are a PLC assistant. If the user is asking a general question, answer directly and concisely.",
                history=history,
            )
            if reply.strip():
                model = str(agent.get("llmModel", "")).strip()
                prefix = f"[{model}] " if model else ""
                return {"ok": True, "message": prefix + reply.strip(), "actions": []}
        except Exception as exc:
            return {"ok": False, "message": f"LLM fallback failed: {type(exc).__name__}: {exc}", "actions": []}

    return {"ok": False, "message": "LLM 未启用或未配置 API Key。请先在 Agent 配置中启用外部 LLM API 并保存。", "actions": []}


def test_llm_config(config):
    agent = normalize_agent(config)
    if not agent.get("llmEnabled", False):
        return {"ok": False, "message": "LLM is disabled."}
    try:
        text = call_openai_responses(agent, "reply with: ok", "You are a test assistant.")
        return {"ok": True, "message": text[:200] or "ok"}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        return {"ok": False, "message": f"HTTP {exc.code}: {detail[:300]}"}
    except Exception as exc:
        return {"ok": False, "message": f"{type(exc).__name__}: {exc}"}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):  # noqa: A003
        return

    def _set_json(self, status=HTTPStatus.OK):
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

    def _set_file(self, content_type, status=HTTPStatus.OK):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

    def _json_body(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {}

    def _send_json(self, payload, status=HTTPStatus.OK):
        self._set_json(status)
        self.wfile.write(json.dumps(payload, ensure_ascii=False).encode("utf-8"))

    def _serve_static(self, path):
        safe = (WEB_DIR / path.lstrip("/")).resolve()
        if WEB_DIR.resolve() not in safe.parents and safe != WEB_DIR.resolve():
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        if not safe.exists() or safe.is_dir():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content_type = "text/plain; charset=utf-8"
        if safe.suffix == ".html":
            content_type = "text/html; charset=utf-8"
        elif safe.suffix == ".css":
            content_type = "text/css; charset=utf-8"
        elif safe.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"
        elif safe.suffix == ".svg":
            content_type = "image/svg+xml"
        elif safe.suffix == ".png":
            content_type = "image/png"
        elif safe.suffix in (".jpg", ".jpeg"):
            content_type = "image/jpeg"
        self._set_file(content_type)
        self.wfile.write(safe.read_bytes())

    def do_GET(self):  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/tools":
            self._send_json({"tools": [build_tool_status(tid) for tid in TOOLS]})
            return
        if parsed.path == "/api/agents":
            self._send_json({"agents": AGENT_STATE["agents"], "currentAgentId": AGENT_STATE["currentAgentId"]})
            return
        if parsed.path in ("/", "/index.html"):
            self._serve_static("index.html")
            return
        if parsed.path in ("/styles.css", "/app.js") or parsed.path.startswith("/logos/"):
            self._serve_static(parsed.path.lstrip("/"))
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):  # noqa: N802
        parsed = urlparse(self.path)
        body = self._json_body()

        if parsed.path == "/api/ai-command":
            agent_id = str(body.get("agentId", "")).strip()
            agent = get_agent(agent_id) if agent_id else get_current_agent()
            if not agent:
                agent = get_current_agent()
            history = body.get("history", [])
            result = ai_handle_command(str(body.get("prompt", "")), agent, history=history)
            result["agentId"] = agent["id"]
            self._send_json(result, status=HTTPStatus.OK if result.get("ok") else HTTPStatus.BAD_REQUEST)
            return

        if parsed.path == "/api/agents/select":
            agent_id = str(body.get("agentId", "")).strip()
            if not get_agent(agent_id):
                self._send_json({"ok": False, "message": "Agent not found."}, status=HTTPStatus.BAD_REQUEST)
                return
            AGENT_STATE["currentAgentId"] = agent_id
            save_agents()
            self._send_json({"ok": True, "currentAgentId": agent_id})
            return

        if parsed.path == "/api/agents/upsert":
            normalized = normalize_agent(body.get("agent", {}))
            existing = get_agent(normalized["id"])
            if existing:
                existing.update(normalized)
            else:
                AGENT_STATE["agents"].append(normalized)
            AGENT_STATE["currentAgentId"] = normalized["id"]
            save_agents()
            self._send_json({"ok": True, "agent": normalized, "currentAgentId": normalized["id"]})
            return

        if parsed.path == "/api/agents/delete":
            agent_id = str(body.get("agentId", "")).strip()
            if agent_id == "default":
                self._send_json({"ok": False, "message": "Default agent cannot be deleted."}, status=HTTPStatus.BAD_REQUEST)
                return
            before = len(AGENT_STATE["agents"])
            AGENT_STATE["agents"] = [item for item in AGENT_STATE["agents"] if item.get("id") != agent_id]
            if len(AGENT_STATE["agents"]) == before:
                self._send_json({"ok": False, "message": "Agent not found."}, status=HTTPStatus.BAD_REQUEST)
                return
            if AGENT_STATE["currentAgentId"] == agent_id:
                AGENT_STATE["currentAgentId"] = AGENT_STATE["agents"][0]["id"]
            save_agents()
            self._send_json({"ok": True, "currentAgentId": AGENT_STATE["currentAgentId"]})
            return

        if parsed.path == "/api/llm/test":
            agent_id = str(body.get("agentId", "")).strip()
            if agent_id:
                agent = get_agent(agent_id)
                if not agent:
                    self._send_json({"ok": False, "message": "Agent not found."}, status=HTTPStatus.BAD_REQUEST)
                    return
                result = test_llm_config(agent)
            else:
                result = test_llm_config(body.get("agent", {}))
            self._send_json(result, status=HTTPStatus.OK if result.get("ok") else HTTPStatus.BAD_REQUEST)
            return

        self.send_error(HTTPStatus.NOT_FOUND)


def main():
    WEB_DIR.mkdir(parents=True, exist_ok=True)
    load_agents()
    with HTTPServer((HOST, PORT), Handler) as server:
        print(f"PLC IDE Controller running at http://{HOST}:{PORT}")
        server.serve_forever()


if __name__ == "__main__":
    main()
