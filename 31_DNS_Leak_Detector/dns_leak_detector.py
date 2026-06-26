"""
30_DNS_Leak_Detector

检测出口 IP 与上游 DNS 解析地是否存在异常偏移，辅助排查 DNS 泄漏风险。
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = BASE / "output"


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


class UniversalDNSLeakDetector:
    def __init__(self, mode: str = "tun", socks_port: int = 10808, timeout_s: int = 10):
        """
        mode:
        - tun: 透明网卡接管模式（不显式注入应用层代理）
        - socks: 强制 requests 走本地 socks5h 代理
        """
        self.mode = mode.lower()
        self.proxy_url = f"socks5h://127.0.0.1:{socks_port}"
        self.timeout_s = timeout_s
        self.session = self._build_robust_session()

    def _build_robust_session(self) -> requests.Session:
        session = requests.Session()
        session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/122.0.0.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Encoding": "gzip, deflate, br",
                "Connection": "keep-alive",
            }
        )

        if self.mode == "socks":
            logger.info("初始化: [SOCKS 模式] 强制流量走本地代理端口。")
            session.proxies = {"http": self.proxy_url, "https": self.proxy_url}
        else:
            logger.info("初始化: [TUN 模式] 信任系统底层路由/虚拟网卡接管。")
            session.trust_env = False

        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET"],
        )
        adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=10, pool_maxsize=10)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        return session

    def verify_core_status(self) -> bool:
        """用 HTTPS Captive 探针做链路探活。"""
        try:
            response = self.session.get("https://www.gstatic.com/generate_204", timeout=5)
            if response.status_code == 204:
                logger.info("链路探活成功 (%s 环境)。", self.mode.upper())
            else:
                logger.warning("探针返回非预期状态码: %s（网络仍可能可用）。", response.status_code)
            return True
        except requests.RequestException as e:
            logger.error("网络阻断: %s", e)
            logger.error("排错: 检查代理客户端路由规则、防火墙、系统代理。")
            return False

    def _fetch_data(self, url: str) -> dict | None:
        try:
            response = self.session.get(url, timeout=self.timeout_s)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            logger.error("接口请求异常 [%s]: %s", url, e)
            return None

    def get_routing_metrics(self) -> tuple[str | None, str | None, dict]:
        diagnostics: dict = {}
        logger.info("正在获取公网出口特征...")
        ip_data = self._fetch_data("http://ip-api.com/json/")
        if not ip_data or ip_data.get("status") != "success":
            return None, None, diagnostics

        proxy_country = ip_data.get("country")
        diagnostics["egress"] = {
            "ip": ip_data.get("query"),
            "country": proxy_country,
            "isp": ip_data.get("isp"),
        }
        logger.info(
            "[公网出口] IP: %s | 节点区: %s | ISP: %s",
            ip_data.get("query"),
            proxy_country,
            ip_data.get("isp"),
        )

        logger.info("正在执行 EDNS 上游回溯...")
        dns_data = self._fetch_data("https://edns.ip-api.com/json")
        if not dns_data or "dns" not in dns_data:
            return proxy_country, None, diagnostics

        dns_country = dns_data["dns"].get("geo")
        diagnostics["dns"] = {
            "ip": dns_data["dns"].get("ip"),
            "country": dns_country,
        }
        logger.info(
            "[上游 DNS] 解析 IP: %s | DNS 区域: %s",
            dns_data["dns"].get("ip"),
            dns_country,
        )
        return proxy_country, dns_country, diagnostics

    @staticmethod
    def _risk_assessment(proxy_country: str, dns_country: str) -> tuple[str, str]:
        if "China" in dns_country and "China" not in proxy_country:
            return (
                "high_risk_dns_leak",
                "高危: 出口为境外但上游 DNS 落在中国，疑似 DNS 泄漏或分流规则异常。",
            )
        if proxy_country not in dns_country and dns_country not in proxy_country:
            return ("possible_anycast_split", "潜在偏移: 出口区与 DNS 区不一致，可能是 Anycast 正常现象。")
        return ("healthy", "状态正常: 出口与 DNS 区域基本对齐。")

    def run_diagnostic(self) -> tuple[bool, dict]:
        logger.info("=" * 55)
        if not self.verify_core_status():
            logger.info("=" * 55)
            return False, {"ok": False, "reason": "connectivity_failed"}

        proxy_country, dns_country, raw = self.get_routing_metrics()
        if not proxy_country or not dns_country:
            logger.error("数据回传失败，无法完成诊断。")
            logger.info("=" * 55)
            return False, {"ok": False, "reason": "metrics_failed", "raw": raw}

        code, msg = self._risk_assessment(proxy_country, dns_country)
        logger.info("-" * 55)
        if code == "high_risk_dns_leak":
            logger.warning("⚠️ %s", msg)
        elif code == "possible_anycast_split":
            logger.info("🔎 %s", msg)
        else:
            logger.info("✅ %s", msg)
        logger.info("=" * 55)

        return True, {
            "ok": True,
            "mode": self.mode,
            "assessment_code": code,
            "assessment_message": msg,
            "raw": raw,
        }


def main() -> None:
    parser = argparse.ArgumentParser(description="30 模块：DNS 泄漏检测（TUN / SOCKS）")
    parser.add_argument("--mode", choices=["tun", "socks"], default="tun", help="检测模式")
    parser.add_argument("--socks-port", type=int, default=10808, help="SOCKS 端口（mode=socks 时生效）")
    parser.add_argument("--timeout", type=int, default=10, help="接口请求超时秒数")
    parser.add_argument("--save-json", action="store_true", help="输出诊断 JSON 到 output/ 目录")
    args = parser.parse_args()

    detector = UniversalDNSLeakDetector(mode=args.mode, socks_port=args.socks_port, timeout_s=args.timeout)
    ok, payload = detector.run_diagnostic()

    if args.save_json:
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = DEFAULT_OUTPUT_DIR / f"dns_diagnostic_{args.mode}_{ts}.json"
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("诊断报告已保存: %s", out)

    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()

