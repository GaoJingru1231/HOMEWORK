"""
最小化测试版 v2 - 显式读取 PORT 环境变量
"""
import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
async def startup():
    port = os.environ.get("PORT", "8000")
    logger.info(f"Server starting on port {port}")

@app.get("/")
async def root():
    port = os.environ.get("PORT", "8000")
    return {"message": "格式快调 API 运行中", "version": "2.0", "port": port}

@app.get("/api/health")
async def health():
    return {"status": "ok"}
