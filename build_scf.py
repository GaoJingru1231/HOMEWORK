"""打包腾讯云 SCF 云函数为可上传的 zip。

用法：
    python build_scf.py

产出：
    dist/format_api.zip  （含 index.py + formatter.py + python-docx 及依赖）

⚠️ 重要：lxml（python-docx 的依赖）含 C 扩展，必须用 Linux 环境打包才能在
   SCF 的 Linux 运行时上运行。请在 WSL / Linux 虚拟机 / GitHub Actions(ubuntu)
   中运行本脚本。在 Windows 上直接打包的 lxml 是 Windows wheel，上传后无法 import。
"""
import sys
import shutil
import subprocess
import zipfile
from pathlib import Path

ROOT = Path(__file__).parent
STAGING = ROOT / "dist" / "format_api"
ZIP_OUT = ROOT / "dist" / "format_api.zip"
FUNC_SRC = ROOT / "cloudfunctions" / "format_api"


def main():
    if STAGING.exists():
        shutil.rmtree(STAGING)
    STAGING.mkdir(parents=True)

    # 1) 源码
    shutil.copy(FUNC_SRC / "index.py", STAGING / "index.py")
    shutil.copy(ROOT / "formatter.py", STAGING / "formatter.py")
    # Web 函数(Type=HTTP)需要 scf_bootstrap 启动脚本 + WSGI 适配器
    shutil.copy(FUNC_SRC / "wsgi_adapter.py", STAGING / "wsgi_adapter.py")
    bootstrap_src = FUNC_SRC / "scf_bootstrap"
    bootstrap_dst = STAGING / "scf_bootstrap"
    shutil.copy(bootstrap_src, bootstrap_dst)
    # zip 内文件权限通过外部 unix 权限位设置(Linux runtime 需要 0o755)
    # zipfile 不直接支持设权限,打包后用 chmod 不可行,改用 tar? — 不,SCF 接受 zip。
    # 实际:SCF 解压时若发现 scf_bootstrap 不可执行会自动 chmod +x,
    # 但保险起见打包时显式设 Unix 可执行位。

    # 2) 依赖（仅 python-docx，避免把 fastapi/uvicorn 打进去）
    #    腾讯云 SCF 最高支持 Python 3.9/3.10，本机是 3.12，必须强制下载 cp39 的
    #    Linux wheel，否则 lxml 的 .so 在 SCF 上无法 import。
    #    用清华镜像源避免国内访问 PyPI 被重置连接。
    subprocess.check_call([
        sys.executable, "-m", "pip", "install",
        "-r", str(FUNC_SRC / "requirements.txt"),
        "--target", str(STAGING),
        "--upgrade",
        "--platform", "manylinux2014_x86_64",
        "--python-version", "39",
        "--implementation", "cp",
        "--abi", "cp39",
        "--only-binary=:all:",
        "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
    ])

    # 3) 打包(为 scf_bootstrap 保留可执行位:Unix external_attr 高位设 0o755)
    ZIP_OUT.parent.mkdir(parents=True, exist_ok=True)
    if ZIP_OUT.exists():
        ZIP_OUT.unlink()
    with zipfile.ZipFile(ZIP_OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in STAGING.rglob("*"):
            if path.is_file():
                arcname = path.relative_to(STAGING).as_posix()
                info = zipfile.ZipInfo.from_file(path, arcname)
                info.compress_type = zipfile.ZIP_DEFLATED
                # scf_bootstrap 在 Linux runtime 需要 0o755
                if arcname == "scf_bootstrap":
                    info.external_attr = (0o755 << 16)
                zf.writestr(info, path.read_bytes())

    size_kb = ZIP_OUT.stat().st_size / 1024
    print(f"打包完成: {ZIP_OUT} ({size_kb:.0f} KB)")
    print("Web 函数(Type=HTTP):含 scf_bootstrap + wsgi_adapter,直接创建函数无需触发器")


if __name__ == "__main__":
    main()
