import os
import sys
import json
import asyncio
from typing import List
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn

# Thêm workspace vào sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from tra_cuu_don_hang import lookup_batch_orders, ensure_headers, export_to_excel

app = FastAPI(title="GHN Batch Order Lookup Web App")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = os.path.join(BASE_DIR, "web_static")
os.makedirs(STATIC_DIR, exist_ok=True)

# Khóa hàng đợi toàn cục: Đảm bảo nhiều đồng nghiệp tra cứu cùng lúc được xếp hàng tuần tự
_lookup_queue_lock = asyncio.Lock()

class LookupRequest(BaseModel):
    order_codes: List[str]

@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
async def serve_home():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h1>Trang tra cứu đang được khởi tạo...</h1>")

@app.get("/api/health")
async def health_check():
    return {
        "status": "ok"
    }

@app.post("/api/lookup")
async def api_lookup(req: LookupRequest):
    if not req.order_codes:
        raise HTTPException(status_code=400, detail="Danh sách mã đơn hàng trống")
    
    # Loại bỏ trùng lặp và làm sạch
    codes = []
    seen = set()
    for c in req.order_codes:
        clean = str(c).strip().upper()
        if clean and clean not in seen:
            seen.add(clean)
            codes.append(clean)
            
    if not codes:
        raise HTTPException(status_code=400, detail="Không có mã đơn hàng hợp lệ")
        
    headers = await ensure_headers()
    if not headers:
        raise HTTPException(status_code=500, detail="Không thể xác thực session GHN. Vui lòng thử lại sau.")
        
    # Xếp hàng yêu cầu tuần tự qua Lock để bảo vệ session GHN không bị nghẽn 429
    async with _lookup_queue_lock:
        results = await lookup_batch_orders(codes, headers)
        
    return {
        "total": len(results),
        "results": results
    }

if __name__ == "__main__":
    print("Khởi chạy Web App Tra Cứu Đơn Hàng GHN tại: http://localhost:8080")
    uvicorn.run(app, host="0.0.0.0", port=8080, log_level="info")
