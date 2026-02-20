from fastapi import FastAPI, Body
from typing import Any, Dict

from app.pipeline.runner201 import run_colab201
from app.pipeline.runner202 import run_colab202
from app.pipeline.runner141 import run_html as run_html141
from app.pipeline.runner142 import run_html as run_html142

app = FastAPI()


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/v1/pipeline")
def pipeline(payload: Dict[str, Any] = Body(...)):
    """
    既存(201/202)と、CF計算書HTML出力(141/142)の両方に対応。

    ■ 201/202 入力例:
      {
        "data": [...],
        "ai_case_id": 21160,
        "loginkey": "...",
        "mode": "201" | "202" | "both"   # 任意（省略時 both）
      }

    ■ 141/142 入力例:
      {
        "mode": "141" | "142" | "both",
        "ai_case_id": 8888,
        "url": "https://...signed xlsx...",
        "s3_bucket": "zlite",
        "s3_region": "ap-northeast-1"
      }
      ※ S3_ACCESS_KEY / S3_SECRET_KEY は環境変数で指定
    """
    mode = str(payload.get("mode", "both")).lower()

    # url がある場合は「HTML生成」モードとして扱う（201/202の互換性を壊さない）
    if payload.get("url"):
        if mode in ("141", "runner141"):
            return {"result": run_html141(payload)}
        if mode in ("142", "runner142"):
            return {"result": run_html142(payload)}
        return {"results": [run_html141(payload), run_html142(payload)]}

