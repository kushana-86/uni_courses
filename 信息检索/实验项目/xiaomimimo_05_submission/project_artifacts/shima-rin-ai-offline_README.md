# Shima Rin AI Offline

这是给树莓派和随身离线设备准备的独立版本。

目标很明确：
- 离线可运行
- 本地 TTS 可控且稳定
- 启动和恢复简单
- 不再依赖实验性的远端语音切换
- 适合带去日本旅行随身使用

## 旅行增强

这版额外加了几组适合日本旅行的离线能力：

- 日本时区默认启用
- 日本旅行打包清单
- 乘车、便利店、温泉等离线注意事项
- 常用旅行日语短句
- 可以直接结合闹钟做赶车、出门、入住提醒

## 现在的语音策略

离线版默认使用 `LOCAL_TTS_BACKEND=auto`：

- 如果本机装了 `piper` 且配置了模型，优先使用 `piper`
- 否则尝试 `espeak-ng`
- 再不行就尝试 `espeak`
- 如果你明确设置为 `browser`，才会使用浏览器语音

这意味着树莓派正式部署时，推荐直接安装本地 TTS，不再依赖浏览器或远端服务。

## 推荐部署方案

### 方案 1：先求稳

使用 `espeak-ng`

优点：
- 安装最简单
- 依赖最少
- 完全离线
- 树莓派上更容易稳定运行

缺点：
- 音色会比较机械
- 不会像角色克隆声线

树莓派示例安装：

```bash
sudo apt update
sudo apt install -y nodejs npm espeak-ng
```

### 方案 2：后续升级

使用 `piper`

优点：
- 比 `espeak-ng` 更自然
- 也能离线运行

缺点：
- 需要额外准备日语模型
- 初次部署比 `espeak-ng` 麻烦一些

## 启动

先复制环境文件：

```bash
cp .env.example .env.local
```

然后启动：

```bash
npm install
npm start
```

默认地址：

```text
http://127.0.0.1:3000
```

## 默认配置

默认就是离线模式：

- `OFFLINE_ONLY_MODE=true`
- `ASSISTANT_REPLY_LANG=ja`
- `ASSISTANT_TIME_ZONE=Asia/Tokyo`
- `ASSISTANT_UTC_OFFSET_HOURS=9`
- `LOCAL_TTS_BACKEND=auto`

也就是说，系统会优先寻找可用的本地 TTS 后端。

## 切换到 Piper

把 `.env.local` 改成类似这样：

```env
OFFLINE_ONLY_MODE=true
ASSISTANT_REPLY_LANG=ja

LOCAL_TTS_BACKEND=piper
LOCAL_TTS_MODEL=./voices/ja_JP-your-voice.onnx
LOCAL_TTS_CONFIG=./voices/ja_JP-your-voice.onnx.json
LOCAL_TTS_SPEAKER=0
LOCAL_TTS_LENGTH_SCALE=1
LOCAL_TTS_NOISE_SCALE=0.667
LOCAL_TTS_NOISE_W=0.8
```

## 只想临时用浏览器语音

把 `.env.local` 改成：

```env
LOCAL_TTS_BACKEND=browser
```

这只建议拿来在电脑上临时测试，不建议作为树莓派正式方案。

## 运行行为

这版已经做了这些稳定化处理：

- 静态资源默认 `no-store`，刷新后不会吃旧脚本缓存
- 前端不会在本地 TTS 失败时偷偷切到另一套语音
- `/api/bootstrap` 会返回当前 TTS 后端状态
- 如果本地语音不可用，页面会直接明确提示原因

## 离线限制

- 当前语音输入仍依赖浏览器的 `SpeechRecognition`
- 在树莓派离线环境里，语音输入不一定稳定
- 文字输入、本地回复、本地播报、记忆和闹钟都可以正常保留

更稳的随身方案是：

- 输入：触屏打字为主
- 输出：本地 TTS 播报

## 建议的树莓派落地步骤

1. 安装 `nodejs`、`npm` 和 `espeak-ng`
2. 把整个 `shima-rin-ai-offline` 文件夹复制到树莓派
3. 运行 `npm install`
4. 运行 `npm start`
5. 打开 `http://127.0.0.1:3000`
6. 确认 `/api/bootstrap` 里 `ttsStatus.available` 为 `true`
7. 再决定要不要继续升级到 `piper`
