"""
最小化测试版 - 验证 Railway 基础部署
"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "格式快调 API 运行中", "version": "1.0.0"}

@app.get("/api/health")
async def health():
    return {"status": "ok"}
