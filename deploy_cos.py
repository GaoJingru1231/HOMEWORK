"""部署前端 index.html 到腾讯云 COS 静态网站。

产出可公网访问的 URL(国内访问快,电脑手机都能开):
    https://<bucket>.cos-website.ap-shanghai.myqcloud.com

用法:
    set TC_SECRET_ID=AKIDxxx
    set TC_SECRET_KEY=xxx
    python deploy_cos.py
"""
import os
import sys
from pathlib import Path
from qcloud_cos import CosConfig, CosS3Client, CosServiceError

# ===== 配置 =====
SECRET_ID = os.environ.get("TC_SECRET_ID", "")
SECRET_KEY = os.environ.get("TC_SECRET_KEY", "")
REGION = "ap-shanghai"
# AppID 从函数 URL 提取(也可在 https://console.cloud.tencent.com/developer 侧边栏看)
APP_ID = "1446811204"
BUCKET = f"fmt-kuaitiao-{APP_ID}"
INDEX_FILE = Path(__file__).parent / "index.html"

if not SECRET_ID or not SECRET_KEY:
    print("ERROR: 请先设置环境变量 TC_SECRET_ID 和 TC_SECRET_KEY")
    sys.exit(1)
if not INDEX_FILE.exists():
    print(f"ERROR: 找不到 {INDEX_FILE}")
    sys.exit(1)

config = CosConfig(
    Region=REGION,
    SecretId=SECRET_ID,
    SecretKey=SECRET_KEY,
    Scheme="https",
)
client = CosS3Client(config)


def ensure_bucket():
    """确保桶存在(不存在则创建,权限公有读私有写)"""
    try:
        client.head_bucket(Bucket=BUCKET)
        print(f"桶已存在: {BUCKET}")
    except CosServiceError as e:
        if e.get_status_code() == 404:
            print(f"创建桶: {BUCKET} ({REGION}, 公有读私有写)")
            client.create_bucket(Bucket=BUCKET, ACL="public-read")
            print("桶创建成功")
        else:
            raise


def upload_index():
    """上传 index.html(覆盖)"""
    print(f"上传 {INDEX_FILE.name} ({INDEX_FILE.stat().st_size} bytes)")
    client.put_object_from_local_file(
        Bucket=BUCKET,
        LocalFilePath=str(INDEX_FILE),
        Key="index.html",
        EnableMD5=True,
    )
    print("上传完成")


def set_static_website():
    """配置静态网站功能(访问根路径自动返回 index.html)"""
    print("配置静态网站...")
    client.put_bucket_website(
        Bucket=BUCKET,
        WebsiteConfiguration={
            "IndexDocument": {"Suffix": "index.html"},
            "ErrorDocument": {"Key": "index.html"},
        },
    )
    print("静态网站已启用")


def print_urls():
    """打印访问 URL"""
    cos_url = f"https://{BUCKET}.cos.{REGION}.myqcloud.com/index.html"
    website_url = f"https://{BUCKET}.cos-website.{REGION}.myqcloud.com/"
    print("\n===== 访问地址 =====")
    print(f"  静态网站域名(推荐,国内快): {website_url}")
    print(f"  COS 直链(备用):            {cos_url}")


def main():
    ensure_bucket()
    upload_index()
    set_static_website()
    print_urls()


if __name__ == "__main__":
    main()
