# PLC 网页控制台

用于本机自动检测并启动 PLC 编程软件，支持可配置 Agent 执行自动化指令（建项目、写程序、粘贴程序）。

## 支持软件

- AutoShop
- GX Works2
- GX Works3
- TIA PortalV18
- InoProShop

## 启动方式

控制台模式（调试）：

```powershell
cd E:\Desktop\AI作业程序\PlcIDE
python .\server.py
```

无控制台模式（推荐）：

- 双击 `start_frontend.vbs`
- 自动后台启动并打开 `http://127.0.0.1:9527`

## Agent 能力

页面内可配置并持久化 Agent：

- Agent 名称、描述
- 默认目标软件
- 执行模式（safe / balanced / aggressive）
- 写程序前自动建项目
- 项目前缀
- 外部 LLM API 参数（Base URL / API Key / Model / Temperature）
- API 连通性测试（测试 API 按钮）

常用指令示例：

- `在 gx works3 创建项目 项目名:TestLine01`
- `在 gx works3 写程序`
- `在 gx works3 粘贴程序`
- `打开 tia portal v18`
- `查看状态`
