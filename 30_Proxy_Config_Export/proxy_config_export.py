"""
代理配置导出工具
自动检测并导出当前系统的代理配置信息
"""
import os
import subprocess
from pathlib import Path
from datetime import datetime

def get_registry_value(key_path, value_name):
    """从Windows注册表读取值"""
    try:
        result = subprocess.run(
            ['reg', 'query', key_path, '/v', value_name],
            capture_output=True,
            text=True,
            encoding='gbk'
        )
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            for line in lines:
                if value_name in line:
                    parts = line.split(None, 2)
                    if len(parts) >= 3:
                        return parts[2].strip()
        return None
    except Exception:
        return None

def get_proxy_config():
    """获取代理配置"""
    ie_settings = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

    # 检查代理是否启用
    proxy_enable = get_registry_value(ie_settings, "ProxyEnable")
    proxy_server = get_registry_value(ie_settings, "ProxyServer")
    proxy_override = get_registry_value(ie_settings, "ProxyOverride")

    # 检查环境变量
    env_http = os.environ.get('HTTP_PROXY') or os.environ.get('http_proxy')
    env_https = os.environ.get('HTTPS_PROXY') or os.environ.get('https_proxy')
    env_no_proxy = os.environ.get('NO_PROXY') or os.environ.get('no_proxy')

    return {
        'proxy_enable': proxy_enable == '0x1',
        'proxy_server': proxy_server,
        'proxy_override': proxy_override,
        'env_http': env_http,
        'env_https': env_https,
        'env_no_proxy': env_no_proxy
    }

def export_proxy_config():
    """导出代理配置到文件"""
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    config = get_proxy_config()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"proxy_config_{timestamp}.txt"

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("代理配置信息导出\n")
        f.write(f"导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")

        # 系统代理设置
        f.write("【系统代理设置】\n")
        f.write("-" * 60 + "\n")
        if config['proxy_enable'] and config['proxy_server']:
            proxy_addr = config['proxy_server']
            f.write("状态: 已启用\n")
            f.write(f"代理服务器: {proxy_addr}\n\n")

            f.write("1. HTTP 代理 (http_proxy)\n")
            f.write(f"   http://{proxy_addr}\n\n")

            f.write("2. HTTPS 代理 (https_proxy)\n")
            f.write(f"   https://{proxy_addr}\n\n")

            if config['proxy_override']:
                f.write("3. 不需要代理的地址 (NO_PROXY)\n")
                no_proxy = config['proxy_override'].replace(';', ',')
                f.write(f"   {no_proxy}\n\n")
        else:
            f.write("状态: 未启用\n\n")

        # 环境变量设置
        f.write("【环境变量设置】\n")
        f.write("-" * 60 + "\n")
        if config['env_http'] or config['env_https']:
            if config['env_http']:
                f.write(f"HTTP_PROXY: {config['env_http']}\n")
            if config['env_https']:
                f.write(f"HTTPS_PROXY: {config['env_https']}\n")
            if config['env_no_proxy']:
                f.write(f"NO_PROXY: {config['env_no_proxy']}\n")
        else:
            f.write("未设置环境变量代理\n")

        f.write("\n" + "=" * 60 + "\n")

    print(f"[OK] 代理配置已导出到: {output_file}")
    return output_file

if __name__ == "__main__":
    export_proxy_config()
