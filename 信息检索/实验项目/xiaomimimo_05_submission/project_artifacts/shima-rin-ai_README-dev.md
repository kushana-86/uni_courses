# 志摩凛 AI 便携原型

这是一个零依赖的最小可行原型，目标是先把下面这条链路跑通:

文字或录音输入 -> 角色回复生成 -> 语音播报 -> 小屏 UI 更新 -> 轻量记忆保存

## 环境要求

- Node.js 20+
- 现代浏览器
  - 推荐 Chrome / Edge
  - 浏览器若支持 `SpeechRecognition`，可以直接录音
  - 浏览器若支持 `speechSynthesis`，可以直接播报回复

## 安装与启动

本项目没有第三方依赖，不需要 `npm install`。

```bash
npm start
```

启动后访问:

```text
http://localhost:3000
```

## 配置 DeepSeek

项目现在会自动读取根目录下的 `.env.local`。你可以先复制一份模板:

```bash
copy .env.example .env.local
```

然后把 `.env.local` 改成这样:

```env
DEEPSEEK_BASE_URL=https://api.deepseek.com
DEEPSEEK_API_KEY=你的真实密钥
DEEPSEEK_MODEL=deepseek-chat
```

说明:

- `.env.local` 已加入 `.gitignore`，不会被默认提交
- 如果同时设置了 `MODEL_API_URL` / `MODEL_API_KEY` / `MODEL_NAME`，项目会优先使用通用兼容接口配置
- 只填 `DEEPSEEK_*` 也可以，服务端会自动拼出聊天接口地址

## 当前结构

```text
.
├─ data
│  ├─ persona-config.json
│  ├─ skill-rules.json
│  └─ user-profile.json
├─ docs
│  └─ shima-rin-persona
├─ public
│  ├─ assets
│  ├─ app.js
│  ├─ index.html
│  └─ styles.css
├─ src
│  ├─ memory-manager.js
│  ├─ persona.js
│  ├─ prompt-builder.js
│  └─ response-engine.js
├─ tests
│  ├─ memory.test.js
│  └─ response-engine.test.js
├─ package.json
└─ server.js
```

## 功能说明

- 角色层
  - `data/persona-config.json` 定义核心角色 Prompt、口头禅和视觉设定
- Skill 路由
  - `data/skill-rules.json` 定义 3 个专属 Skill 的触发词和执行规则
  - 当前支持 `露营规划`、`治愈陪伴`、`户外知识解答`
- 轻量记忆
  - `data/user-profile.json` 本地保存用户名字、偏好、计划、重要日期和聊天风格
- 回复生成
  - 默认用本地规则回复器
  - 如果配置远端模型，会优先走远端模型，失败时自动回退本地模式
- 语音链路
  - 浏览器负责录音识别与 TTS
  - 服务端负责角色逻辑与记忆注入
- UI
  - 面向 5 到 7 英寸小屏
  - 状态包括待机、聆听、思考、回答中
  - 会显示当前触发的 Skill

## 远端模型可选配置

如果你有兼容 OpenAI 风格的接口，也可以直接设置下面的环境变量:

```powershell
$env:MODEL_API_URL="https://your-endpoint.example.com/v1/chat/completions"
$env:MODEL_API_KEY="your-key"
$env:MODEL_NAME="your-model"
npm start
```

说明:

- 未配置时，项目使用本地回复模式
- 已配置时，服务端会把角色提示词和相关记忆一起注入
- 如果远端失败，会自动回退本地模式
- 已提供 `.env.local` 自动加载，不一定要每次手动设 PowerShell 环境变量

## 常见问题

### 1. 点了“开始录音”没反应

最常见原因:

- 浏览器不支持 `SpeechRecognition`
- 页面不是在受支持环境里打开
- 麦克风权限没有允许

排查:

- 优先用最新版 Edge 或 Chrome
- 检查地址是否为 `http://localhost:3000`
- 检查浏览器麦克风权限

### 2. 没有语音播报

最常见原因:

- 浏览器禁用了 `speechSynthesis`
- 系统没有可用语音

排查:

- 换 Edge / Chrome
- 点击“停止播报”后再试一次

### 3. 回复比较简单

这是当前最小原型的预期表现。未配置远端模型时，回复由本地规则引擎生成，但已经会按 Skill 路由切换到露营规划、治愈陪伴或户外知识模式。

### 4. 记忆好像没更新

检查:

- `data/user-profile.json`
- 输入里是否明确说到了名字、喜欢什么、计划、日期或聊天偏好

## 验证方式

1. 启动 `npm start`
2. 打开 `http://localhost:3000`
3. 输入:
   - `我叫小林`
   - `我喜欢热咖啡`
   - `下周末想去露营`
   - `帮我规划一次露营`
   - `陪我说说话`
   - `户外怎么保暖`
4. 检查右下方记忆面板是否更新
5. 检查界面上的 `当前 Skill` 是否随输入变化
6. 点击“开始录音”测试语音输入
7. 观察状态是否按 `待机 -> 聆听中 -> 思考中 -> 回答中 -> 待机` 切换

## 后续可扩展方向

- 接入真实 LLM
- 增加天气 API
- 增加多角色表情资源
- 增加低性能设备模式
- 迁移到树莓派小屏
