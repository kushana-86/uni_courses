import json
import sys
from pathlib import Path

from rapidocr_onnxruntime import RapidOCR

start = int(sys.argv[1])
end = int(sys.argv[2])
engine = RapidOCR()
folder = Path("data/2024B_images")
out_dir = Path("data/2024B_ocr_parts")
out_dir.mkdir(exist_ok=True)

for number in range(start, end):
    source = folder / f"{number:03}.png"
    target = out_dir / f"{number:03}.json"
    if target.exists() or not source.exists():
        continue
    result, _ = engine(str(source))
    target.write_text(json.dumps(result or [], ensure_ascii=False), encoding="utf-8")
    print(number, len(result or []))
