"""部署 format_api 到腾讯云 SCF，并配置 HTTP 触发器，打印公网 URL。

用法：
    python deploy_scf.py
"""
import os
import sys
import json
import base64
import time
from pathlib import Path

from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.scf.v20180416 import scf_client, models

# ===== 配置（从环境变量读，避免硬编码）=====
SECRET_ID = os.environ.get("TC_SECRET_ID", "")
SECRET_KEY = os.environ.get("TC_SECRET_KEY", "")
REGION = os.environ.get("TC_REGION", "ap-shanghai")
FUNCTION_NAME = "format_api"
NAMESPACE = "default"
ZIP_PATH = Path(__file__).parent / "dist" / "format_api.zip"

if not SECRET_ID or not SECRET_KEY:
    print("ERROR: 请先设置环境变量 TC_SECRET_ID 和 TC_SECRET_KEY")
    sys.exit(1)
if not ZIP_PATH.exists():
    print(f"ERROR: zip 不存在: {ZIP_PATH}，先运行 python build_scf.py")
    sys.exit(1)

cred = credential.Credential(SECRET_ID, SECRET_KEY)
http_prof = HttpProfile(endpoint="scf.tencentcloudapi.com", reqTimeout=120)
prof = ClientProfile(httpProfile=http_prof)
client = scf_client.ScfClient(cred, REGION, prof)


def del_if_exists():
    """删除已存在的同名函数（重新部署用）"""
    try:
        req = models.DeleteFunctionRequest()
        req.FunctionName = FUNCTION_NAME
        req.Namespace = NAMESPACE
        client.DeleteFunction(req)
        print(f"已删除旧函数 {FUNCTION_NAME}")
        time.sleep(3)
    except Exception as e:
        # 不存在就忽略
        msg = str(e)
        if "ResourceNotFound" in msg or "not found" in msg.lower() or "404" in msg:
            print(f"无旧函数，直接创建")
        else:
            print(f"删除旧函数时警告（忽略继续）: {msg[:120]}")


def create_function():
    zip_b64 = base64.b64encode(ZIP_PATH.read_bytes()).decode("ascii")
    req = models.CreateFunctionRequest()
    req.FunctionName = FUNCTION_NAME
    req.Namespace = NAMESPACE
    req.Runtime = "Python3.9"
    req.Handler = "index.main_handler"
    req.MemorySize = 512
    req.Timeout = 60
    req.Code = models.Code()
    req.Code.ZipFile = zip_b64
    req.Type = "HTTP"  # Web 函数:自带公网 HTTP 入口,无需 API 网关触发器(网关已停售)
    req.Description = "格式快调 - Word 文档排版 API"

    print(f"正在创建 Web 函数（zip {ZIP_PATH.stat().st_size/1024:.0f}KB, Python3.9, 512MB, 60s）...")
    resp = client.CreateFunction(req)
    print(f"创建请求已提交: {resp.RequestId}")
    return resp


def wait_ready(timeout=120):
    """轮询函数状态直到 Active"""
    print("等待函数就绪...", end="", flush=True)
    start = time.time()
    while time.time() - start < timeout:
        try:
            req = models.GetFunctionRequest()
            req.FunctionName = FUNCTION_NAME
            req.Namespace = NAMESPACE
            r = client.GetFunction(req)
            status = r.Status
            if status == "Active":
                print(f" Active")
                return r
            print(f" {status}", end="", flush=True)
        except Exception as e:
            print(f"?", end="", flush=True)
        time.sleep(3)
    raise RuntimeError("函数长时间未就绪")


def create_http_trigger():
    """创建 API 网关触发器,获取公网 URL。

    腾讯云 SCF 的 HTTP 触发器实际是 API 网关触发器,TriggerDesc 必须用
    API 网关格式(字段 authRequired/isIntegratedResponse 为字符串 "TRUE"/"FALSE",
    不是老版 AuthType/IntegratedResponse 布尔值)。
    参考: https://cloud.tencent.com/document/product/583/39901
    """
    req = models.CreateTriggerRequest()
    req.FunctionName = FUNCTION_NAME
    req.Namespace = NAMESPACE
    req.TriggerName = "api_trigger"
    req.Type = "apigw"
    req.TriggerDesc = json.dumps({
        "api": {
            "authRequired": "FALSE",
            "requestConfig": {"method": "ANY"},
            "isIntegratedResponse": "TRUE",
        },
        "service": {"serviceName": "SCF_API_SERVICE"},
        "release": {"environmentName": "release"},
    }, separators=(",", ":"))  # 腾讯云要求 JSON 连续无空格
    print("正在创建 API 网关触发器...")
    try:
        resp = client.CreateTrigger(req)
        print(f"触发器创建: {resp.RequestId}")
        return resp
    except Exception as e:
        msg = str(e)
        if "already" in msg.lower() or "exist" in msg.lower():
            print("触发器已存在,跳过")
        else:
            raise


def list_triggers():
    """查询触发器拿到 URL"""
    req = models.ListTriggersRequest()
    req.FunctionName = FUNCTION_NAME
    req.Namespace = NAMESPACE
    req.Limit = 20
    resp = client.ListTriggers(req)
    triggers = resp.Triggers or []
    return triggers


def main():
    del_if_exists()
    create_function()
    wait_ready()
    # Web 函数(Type=HTTP)自带公网入口,无需创建触发器
    # API 网关触发器已于 2024 年停止售卖,改用 Web 函数原生 HTTP 入口
    base = f"https://{REGION}.scf.tencentcloudapi.com/{NAMESPACE}/{FUNCTION_NAME}"
    print("\n===== 公网访问地址 =====")
    print(f"  {base}")
    print(f"\n健康检查: {base}/api/health")
    print(f"模板列表: {base}/api/templates")


if __name__ == "__main__":
    main()
