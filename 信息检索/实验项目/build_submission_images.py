from pathlib import Path
import json
import textwrap

from PIL import Image, ImageDraw, ImageFont


ROOT = Path("xiaomimimo_05_submission")
LOG_DIR = ROOT / "02_agent_workflow_logs"
LINK_DIR = ROOT / "03_github_links_demo"


def load_font(size=28, bold=False):
    candidates = [
        Path("C:/Windows/Fonts/msyhbd.ttc") if bold else Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def wrap_line(line, width=54):
    chunks = []
    current = ""
    visual = 0
    for ch in line:
        step = 2 if ord(ch) > 127 else 1
        if visual + step > width:
            chunks.append(current)
            current = ch
            visual = step
        else:
            current += ch
            visual += step
    if current:
        chunks.append(current)
    return chunks or [""]


def draw_text_image(title, sections, output, width=1400):
    title_font = load_font(38, True)
    h_font = load_font(28, True)
    body_font = load_font(25, False)
    small_font = load_font(21, False)

    lines = []
    for heading, body in sections:
        lines.append(("heading", heading))
        for raw in str(body).splitlines():
            for wrapped in wrap_line(raw, 62):
                lines.append(("body", wrapped))
        lines.append(("space", ""))

    height = 150 + len(lines) * 38 + 60
    image = Image.new("RGB", (width, height), "#fbfbf7")
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width, 92], fill="#163b34")
    draw.text((48, 24), title, fill="white", font=title_font)

    y = 125
    for kind, text in lines:
        if kind == "heading":
            draw.text((56, y), text, fill="#123c35", font=h_font)
            y += 42
        elif kind == "body":
            draw.text((78, y), text, fill="#26312f", font=body_font)
            y += 34
        else:
            y += 14

    draw.text((56, height - 42), "材料包自动生成，仅作申请证明辅助。", fill="#66736f", font=small_font)
    image.save(output)


def main():
    chat_path = LOG_DIR / "shima-rin-ai_chat_workflow_response.json"
    bootstrap_path = LOG_DIR / "shima-rin-ai_bootstrap_response.json"
    models_path = LOG_DIR / "shima-rin-ai_live2d_models_response.json"

    chat = json.loads(chat_path.read_text(encoding="utf-8-sig")) if chat_path.exists() else {}
    bootstrap = json.loads(bootstrap_path.read_text(encoding="utf-8-sig")) if bootstrap_path.exists() else {}
    models = json.loads(models_path.read_text(encoding="utf-8-sig")) if models_path.exists() else []

    chat_sections = [
        ("项目", "志摩凛 AI 便携陪伴 Agent"),
        ("本地演示", "E:/shima-rin-ai\nhttp://127.0.0.1:3000"),
        ("输入样例", "我叫周同学，我下周想去露营，帮我规划一次露营"),
        ("Agent 返回", json.dumps(chat, ensure_ascii=False, indent=2)[:1500]),
        ("初始化能力", f"appName: {bootstrap.get('appName', '')}\nmodels: {len(models)} 个 Live2D/3D 模型配置\nskills: 露营规划、治愈陪伴、户外知识解答、时间/闹钟工具"),
    ]
    draw_text_image("志摩凛 AI Agent 工作流证明", chat_sections, LOG_DIR / "shima-rin-ai_workflow_proof.png")

    link_sections = [
        ("志摩凛 AI", "本地项目：E:/shima-rin-ai\n离线版：E:/shima-rin-ai-offline\n演示地址：http://127.0.0.1:3000"),
        ("AI/信息分析 INF_A", "https://github.com/kushana-86/INF_A"),
        ("微信小程序 择域先知", "https://github.com/kushana-86/zyxz-tibet-job-helper-miniapp"),
        ("文明六决策系统", "https://github.com/kushana-86/civ6-city-decision-system"),
        ("语义重排序实验", "本地目录：C:/Users/90513/Desktop/jiansuo/XXXXXX_周同学"),
    ]
    draw_text_image("GitHub 项目链接与演示地址", link_sections, LINK_DIR / "github_links_and_demo_addresses.png")

    billing_sections = [
        ("状态", "本机未找到可直接提交的过去 30 天 AI 平台账单截图。"),
        ("请补充", "请将 OpenAI / DeepSeek / Claude / Kimi / 通义千问 / 智谱等平台的真实账单截图放入 01_ai_billing 文件夹。"),
        ("建议命名", "openai_billing_last_30_days.png\ndeepseek_billing_last_30_days.png\nclaude_billing_last_30_days.png"),
        ("说明", "不要上传伪造账单；没有账单时，优先上传 Agent 工作流截图、终端日志和 GitHub 项目链接。"),
    ]
    draw_text_image("AI 平台账单截图待补充说明", billing_sections, ROOT / "01_ai_billing" / "AI平台账单截图_待补充说明.png")


if __name__ == "__main__":
    main()
