"""
WiFi 密码查看工具（13 模块）

说明：
- 仅在 Windows 环境可用（使用 netsh）
- 需具备读取 WiFi 配置权限（部分设备需管理员权限）

默认目录：
- 输出：25_WiFi_Passwords/output/wifi_passwords.csv

用法：
  python wifi_passwords.py
  python wifi_passwords.py --output "output/wifi_passwords.csv"
  python wifi_passwords.py --encoding gbk --quiet

"""
from __future__ import annotations

import argparse
import csv
import subprocess
from pathlib import Path
from collections.abc import Iterable



PROFILE_MARKERS = ("All User Profile", "所有用户配置文件")
KEY_MARKERS = ("Key Content", "关键内容")


def run_netsh(args: list[str], encoding: str) -> str:
    try:
        return subprocess.check_output(args).decode(encoding, errors="ignore")
    except FileNotFoundError:
        print("错误：未找到 netsh，请确认在 Windows 环境运行。")
        return ""
    except subprocess.CalledProcessError as exc:
        print(f"错误：netsh 执行失败（{exc.returncode}）。可能需要管理员权限。")
        return ""



def extract_profiles(output: str) -> list[str]:
    profiles: list[str] = []
    for line in output.splitlines():
        if any(marker in line for marker in PROFILE_MARKERS):
            parts = line.split(":", 1)
            if len(parts) == 2:
                name = parts[1].strip()
                if name:
                    profiles.append(name)
    return profiles


def extract_password(output: str) -> str:
    for line in output.splitlines():
        if any(marker in line for marker in KEY_MARKERS):
            parts = line.split(":", 1)
            if len(parts) == 2:
                return parts[1].strip()
    return ""


def iter_wifi_credentials(encoding: str) -> Iterable[tuple[str, str]]:
    profiles_output = run_netsh(["netsh", "wlan", "show", "profiles"], encoding)
    profiles = extract_profiles(profiles_output)
    for name in profiles:
        profile_output = run_netsh(
            ["netsh", "wlan", "show", "profile", name, "key=clear"],
            encoding,
        )
        password = extract_password(profile_output)
        if password:
            yield name, password



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="WiFi 密码查看工具（Windows）")
    parser.add_argument("--output", default=None, help="输出 CSV 路径")
    parser.add_argument("--encoding", default="gbk", help="netsh 输出编码（默认 gbk）")
    parser.add_argument("--quiet", action="store_true", help="不在控制台显示")
    return parser.parse_args()



def main() -> None:
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    args = parse_args()
    output_path = Path(args.output) if args.output else output_dir / "wifi_passwords.csv"
    if not output_path.is_absolute():
        output_path = output_dir / output_path.name

    rows = list(iter_wifi_credentials(args.encoding))

    with output_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ssid", "password"])
        for ssid, password in rows:
            writer.writerow([ssid, password])

    if not args.quiet:
        for ssid, password in rows:
            print(f"网络名称：{ssid}, 密码：{password}")

    print(f"已输出：{output_path}（共 {len(rows)} 条）")



if __name__ == "__main__":
    main()
