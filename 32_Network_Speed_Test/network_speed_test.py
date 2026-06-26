"""
31_Network_Speed_Test

实时测速：局域网（网关延迟/可选内网 URL）、国内网、国外网、VPN 实际吞吐（直连 vs SOCKS 对比）。
局域网占用排查：扫描在线设备 IP/MAC，并指引在路由器上按 IP/MAC 限速（单机无法直接测他人实时带宽）。
菜单 4 为网速慢、查「谁占网」时使用；不提供对路由器首页的 HTTP 压测（无测速文件时无意义）。
"""

from __future__ import annotations

import argparse
import ipaddress
import json
import logging
import platform
import re
import socket
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
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
logging.getLogger("urllib3.connectionpool").setLevel(logging.ERROR)

CHUNK_SIZE = 64 * 1024
SCAN_PING_WORKERS = 64
PORT_SCAN_TIMEOUT = 0.25
PING_TIMEOUT_MS_WIN = 200
HOSTNAME_RESOLVE_LIMIT = 12
FAST_PING_CAP = 48
PORTS_QUICK = (443, 80, 445)


def _win_no_window_flags() -> int:
    if platform.system() == "Windows":
        return getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return 0


def _run_hidden(cmd: list[str], timeout: int = 10, text: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        capture_output=True,
        text=text,
        encoding="gbk" if platform.system() == "Windows" else "utf-8",
        errors="replace",
        timeout=timeout,
        creationflags=_win_no_window_flags(),
    )


def _check_output_hidden(cmd: list[str], timeout: int = 10) -> str:
    return subprocess.check_output(
        cmd,
        text=True,
        encoding="gbk" if platform.system() == "Windows" else "utf-8",
        errors="replace",
        timeout=timeout,
        stderr=subprocess.DEVNULL,
        creationflags=_win_no_window_flags(),
    )

# (名称, URL, 最大下载字节；0 表示仅测延迟)
DOMESTIC_PROBES = [
    ("百度", "https://www.baidu.com/img/flexible/logo/pc/result.png", 512_000),
    ("阿里云镜像", "https://mirrors.aliyun.com/debian/dists/stable/Release", 256_000),
    ("网易", "https://www.163.com/favicon.ico", 128_000),
]

INTERNATIONAL_PROBES = [
    ("Cloudflare", "https://speed.cloudflare.com/__down?bytes=10485760", 10_485_760),
    ("Google", "https://www.gstatic.com/generate_204", 0),
    ("Microsoft", "https://www.microsoft.com/favicon.ico", 256_000),
]

VPN_COMPARE_URL = "https://speed.cloudflare.com/__down?bytes=5242880"

CATEGORY_TITLES = {
    "lan": "局域网",
    "domestic": "国内网",
    "international": "国外网",
    "vpn": "VPN 对比",
}


@dataclass
class SpeedSample:
    label: str
    category: str
    route: str
    url: str = ""
    latency_ms: float | None = None
    download_mbps: float | None = None
    bytes_downloaded: int = 0
    duration_s: float = 0.0
    ok: bool = True
    error: str = ""
    probe_type: str = "download"  # download | latency | ping

    @property
    def short_name(self) -> str:
        if "/" in self.label:
            return self.label.split("/", 1)[1]
        return self.label


@dataclass
class RunReport:
    timestamp: str
    platform: str
    gateway: str = ""
    local_ip: str = ""
    socks_ready: bool = False
    samples: list[SpeedSample] = field(default_factory=list)
    vpn_summary: dict[str, object] = field(default_factory=dict)


def _socks_ready() -> tuple[bool, str]:
    try:
        import socks  # noqa: F401  # PySocks

        return True, ""
    except ImportError:
        return False, "未安装 PySocks（pip install PySocks）"


def _default_gateway() -> str:
    system = platform.system()
    try:
        if system == "Windows":
            out = subprocess.check_output(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    "(Get-NetRoute -DestinationPrefix '0.0.0.0/0' | "
                    "Sort-Object RouteMetric | Select-Object -First 1).NextHop",
                ],
                text=True,
                timeout=8,
                stderr=subprocess.DEVNULL,
            )
            gw = out.strip()
            if gw:
                return gw
            out = subprocess.check_output("route print 0.0.0.0", shell=True, text=True, timeout=8)
            for line in out.splitlines():
                parts = line.split()
                if len(parts) >= 3 and parts[0] == "0.0.0.0":
                    return parts[2]
        else:
            out = subprocess.check_output(["ip", "route"], text=True, timeout=8)
            m = re.search(r"default via (\S+)", out)
            if m:
                return m.group(1)
    except (subprocess.SubprocessError, OSError, ValueError) as e:
        logger.debug("默认网关解析失败: %s", e)
    return ""


def _ping_ms(host: str, count: int = 4) -> tuple[float | None, float | None, str]:
    """返回 (平均 ms, 丢包率 0~100, 错误说明)。"""
    if not host:
        return None, None, "无网关地址"
    flag = "-n" if platform.system() == "Windows" else "-c"
    try:
        out = _check_output_hidden(["ping", host, flag, str(count)], timeout=20)
    except (subprocess.SubprocessError, OSError) as e:
        return None, None, str(e)

    times: list[float] = []
    for line in out.splitlines():
        for pat in (
            r"(?:time|时间)\s*[=<]\s*(\d+(?:\.\d+)?)\s*ms",
            r"(?:time|时间)\s*[=<](\d+(?:\.\d+)?)\s*ms",
            r"(\d+(?:\.\d+)?)\s*ms\s+TTL",
        ):
            m = re.search(pat, line, re.I)
            if m:
                times.append(float(m.group(1)))
                break

    loss_pct: float | None = None
    m_loss = re.search(r"(\d+)\s*%\s*(?:loss|丢失)", out, re.I)
    if m_loss:
        loss_pct = float(m_loss.group(1))

    if not times:
        for pat in (
            r"(?:Average|平均)\s*=\s*(\d+(?:\.\d+)?)\s*ms",
            r"平均\s*=\s*(\d+(?:\.\d+)?)\s*ms",
        ):
            m = re.search(pat, out, re.I)
            if m:
                return float(m.group(1)), loss_pct, ""

    if times:
        return round(sum(times) / len(times), 2), loss_pct, ""
    return None, loss_pct, "未解析到 RTT（可尝试管理员权限或检查防火墙）"


def _build_session(route: str, socks_port: int) -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
        }
    )
    route = route.lower()
    if route == "socks":
        proxy = f"socks5h://127.0.0.1:{socks_port}"
        session.proxies = {"http": proxy, "https": proxy}
    else:
        session.trust_env = False

    retry = Retry(total=2, backoff_factor=0.5, status_forcelist=[502, 503, 504], allowed_methods=["GET", "HEAD"])
    adapter = HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def _request_latency(session: requests.Session, url: str, timeout_s: int) -> float | None:
    t0 = time.perf_counter()
    try:
        with session.get(url, timeout=timeout_s, stream=True) as resp:
            resp.raise_for_status()
            next(resp.iter_content(256), None)
    except requests.RequestException:
        return None
    return round((time.perf_counter() - t0) * 1000.0, 2)


def _download_speed(
    session: requests.Session,
    url: str,
    max_bytes: int,
    timeout_s: int,
    label: str,
    live: bool,
) -> tuple[int, float]:
    started = time.perf_counter()
    downloaded = 0
    last_print = started

    with session.get(url, timeout=timeout_s, stream=True) as resp:
        resp.raise_for_status()
        for chunk in resp.iter_content(chunk_size=CHUNK_SIZE):
            if not chunk:
                continue
            downloaded += len(chunk)
            now = time.perf_counter()
            if live and now - last_print >= 0.5:
                elapsed = now - started
                if elapsed > 0:
                    mbps = (downloaded * 8) / elapsed / 1_000_000
                    logger.info("  [%s] 进行中 %.2f Mbps (%s)", label, mbps, _human_bytes(downloaded))
                last_print = now
            if downloaded >= max_bytes:
                break

    return downloaded, time.perf_counter() - started


def _human_bytes(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024 ** 2:
        return f"{n / 1024:.1f} KB"
    return f"{n / 1024 ** 2:.2f} MB"


def _mbps(byte_count: int, seconds: float) -> float | None:
    if seconds <= 0 or byte_count <= 0:
        return None
    return round((byte_count * 8) / seconds / 1_000_000, 2)


def _status_icon(ok: bool) -> str:
    return "✓" if ok else "✗"


def _format_metric(sample: SpeedSample) -> str:
    if sample.probe_type == "ping" and sample.latency_ms is not None:
        text = f"{sample.latency_ms} ms"
        if sample.error and "丢失" in sample.error:
            text += f" ({sample.error})"
        return text
    if sample.probe_type == "latency" and sample.latency_ms is not None:
        return f"延迟 {sample.latency_ms} ms"
    if sample.download_mbps is not None:
        return f"{sample.download_mbps:.2f} Mbps"
    if sample.latency_ms is not None:
        return f"延迟 {sample.latency_ms} ms"
    return "—"


def _best_download(samples: list[SpeedSample]) -> SpeedSample | None:
    ok_dl = [s for s in samples if s.ok and s.download_mbps is not None]
    if not ok_dl:
        return None
    return max(ok_dl, key=lambda s: s.download_mbps or 0)


def _avg_download_mbps(samples: list[SpeedSample]) -> float | None:
    vals = [s.download_mbps for s in samples if s.ok and s.download_mbps is not None]
    if not vals:
        return None
    return round(sum(vals) / len(vals), 2)


def print_summary(report: RunReport) -> None:
    """终端输出结构化汇总（测速结束后阅读此块即可）。"""
    w = 58
    lines: list[str] = []
    lines.append("")
    lines.append("=" * w)
    lines.append("  测速结果汇总".center(w))
    lines.append("=" * w)

    by_cat: dict[str, list[SpeedSample]] = {}
    for s in report.samples:
        by_cat.setdefault(s.category, []).append(s)

    # 局域网
    lines.append("")
    lines.append("【局域网】")
    lan = by_cat.get("lan", [])
    ping_s = next((s for s in lan if s.probe_type == "ping"), None)
    if ping_s and ping_s.ok:
        gw = report.gateway or ping_s.url
        extra = f"，丢包 {ping_s.error}" if ping_s.error and "%" in ping_s.error else ""
        lines.append(f"  默认网关 {gw}  Ping  {_format_metric(ping_s)}{extra}  {_status_icon(True)}")
    elif ping_s:
        lines.append(f"  网关 Ping  失败（{ping_s.error or '未知'}）  {_status_icon(False)}")
    else:
        lines.append("  未测网关 Ping")
    if report.local_ip:
        lines.append(f"  本机 IP     {report.local_ip}")
    for s in lan:
        if s.probe_type == "download":
            lines.append(f"  内网 HTTP   {s.short_name}  {_format_metric(s)}  {_status_icon(s.ok)}")

    # 国内 / 国外
    for cat in ("domestic", "international"):
        samples = by_cat.get(cat, [])
        if not samples:
            continue
        title = CATEGORY_TITLES[cat]
        route = samples[0].route.upper() if samples else "—"
        lines.append("")
        lines.append(f"【{title}】（{route}）")
        best = _best_download(samples)
        avg = _avg_download_mbps(samples)
        if best:
            lines.append(f"  最快        {best.short_name}  {best.download_mbps:.2f} Mbps  {_status_icon(True)}")
        if avg is not None:
            lines.append(f"  下载平均    {avg:.2f} Mbps")
        for s in samples:
            if s.probe_type == "latency":
                lines.append(f"  探活        {s.short_name}  {_format_metric(s)}  {_status_icon(s.ok)}")
            elif s is not best:
                mark = _status_icon(s.ok)
                detail = _format_metric(s) if s.ok else (s.error[:40] if s.error else "失败")
                lines.append(f"  明细        {s.short_name}  {detail}  {mark}")

    # VPN
    vpn_samples = by_cat.get("vpn", [])
    if vpn_samples:
        lines.append("")
        lines.append("【VPN 实际网速】")
        direct = next((s for s in vpn_samples if "直连" in s.short_name), None)
        socks = next((s for s in vpn_samples if "SOCKS" in s.short_name), None)
        if direct and direct.ok:
            lines.append(f"  系统路由    {direct.download_mbps:.2f} Mbps  {_status_icon(True)}")
        elif direct:
            lines.append(f"  系统路由    失败  {_status_icon(False)}")
        if socks and socks.ok:
            lines.append(f"  SOCKS 代理  {socks.download_mbps:.2f} Mbps  {_status_icon(True)}")
        elif socks:
            err = socks.error or "失败"
            if "SOCKS support" in err or "PySocks" in err:
                err = "缺少 PySocks，请执行: pip install PySocks"
            lines.append(f"  SOCKS 代理  {err}  {_status_icon(False)}")

        conclusion = str(report.vpn_summary.get("conclusion", ""))
        if conclusion:
            lines.append(f"  结论        {conclusion}")

    lines.append("")
    lines.append("=" * w)
    lines.append("")

    for line in lines:
        logger.info(line)


class NetworkSpeedTester:
    def __init__(
        self,
        socks_port: int = 10808,
        timeout_s: int = 15,
        max_bytes: int = 10_485_760,
        live: bool = True,
        lan_url: str = "",
    ):
        self.socks_port = socks_port
        self.timeout_s = timeout_s
        self.max_bytes = max_bytes
        self.live = live
        self.lan_url = lan_url.strip()
        self.gateway = _default_gateway()
        self.socks_ready, self.socks_hint = _socks_ready()

    def _probe(
        self,
        session: requests.Session,
        route: str,
        category: str,
        name: str,
        url: str,
        cap_bytes: int,
    ) -> SpeedSample:
        label = f"{category}/{name}"
        sample = SpeedSample(label=label, category=category, route=route, url=url)

        if cap_bytes == 0:
            sample.probe_type = "latency"
            lat = _request_latency(session, url, self.timeout_s)
            sample.latency_ms = lat
            sample.ok = lat is not None
            if not sample.ok:
                sample.error = "探活失败"
            logger.info("  %s | %s %s", label, _format_metric(sample), _status_icon(sample.ok))
            return sample

        cap = min(cap_bytes, self.max_bytes)
        sample.probe_type = "download"
        try:
            n, sec = _download_speed(session, url, cap, self.timeout_s, label, self.live)
            sample.bytes_downloaded = n
            sample.duration_s = round(sec, 3)
            sample.download_mbps = _mbps(n, sec)
            sample.ok = n > 0 and sec > 0
            logger.info(
                "  %s | %.2f Mbps | %s / %.2fs %s",
                label,
                sample.download_mbps or 0,
                _human_bytes(n),
                sec,
                _status_icon(sample.ok),
            )
        except requests.RequestException as e:
            sample.ok = False
            sample.error = str(e)
            logger.warning("  %s | 失败: %s", label, e)
        return sample

    def run_lan(self, report: RunReport) -> None:
        logger.info("—— 局域网 ——")
        if self.gateway:
            logger.info("默认网关: %s", self.gateway)
            avg_ms, loss, err = _ping_ms(self.gateway)
            loss_note = f"丢包 {loss:g}%" if loss is not None else ""
            sample = SpeedSample(
                label="LAN/默认网关 Ping",
                category="lan",
                route="local",
                url=self.gateway,
                latency_ms=avg_ms,
                ok=avg_ms is not None,
                error=loss_note or err,
                probe_type="ping",
            )
            report.samples.append(sample)
            if sample.ok:
                logger.info("  网关 Ping | %.2f ms %s %s", avg_ms, loss_note, _status_icon(True))
            else:
                logger.info("  网关 Ping | 失败 (%s) %s", err, _status_icon(False))
        else:
            logger.warning("  未能解析默认网关，跳过 Ping。")

        if self.lan_url:
            session = _build_session("direct", self.socks_port)
            report.samples.append(
                self._probe(session, "direct", "lan", "内网 HTTP", self.lan_url, self.max_bytes)
            )

        try:
            report.local_ip = socket.gethostbyname(socket.gethostname())
            logger.info("  本机 %s | IP %s", socket.gethostname(), report.local_ip)
        except OSError:
            pass

    def run_category(
        self,
        report: RunReport,
        title: str,
        category: str,
        probes: list[tuple[str, str, int]],
        route: str,
    ) -> None:
        logger.info("—— %s（%s）——", title, route.upper())
        if route == "socks" and not self.socks_ready:
            logger.warning("  跳过：%s", self.socks_hint)
            return
        if route == "socks":
            logger.info("  路由: SOCKS5h 127.0.0.1:%s", self.socks_port)
        else:
            logger.info("  路由: 系统直连（TUN/物理路由，trust_env=False）")
        session = _build_session(route, self.socks_port)
        for name, url, cap in probes:
            report.samples.append(self._probe(session, route, category, name, url, cap))

    def run_vpn_compare(self, report: RunReport) -> None:
        logger.info("—— VPN 实际网速（同 URL：直连 vs SOCKS）——")
        cap = min(5_242_880, self.max_bytes)

        logger.info("  路由: 系统直连")
        direct_session = _build_session("direct", self.socks_port)
        direct = self._probe(direct_session, "direct", "vpn", "Cloudflare-直连", VPN_COMPARE_URL, cap)

        socks: SpeedSample
        if not self.socks_ready:
            socks = SpeedSample(
                label="vpn/Cloudflare-SOCKS",
                category="vpn",
                route="socks",
                url=VPN_COMPARE_URL,
                ok=False,
                error=self.socks_hint,
            )
            logger.warning("  vpn/Cloudflare-SOCKS | 跳过: %s", self.socks_hint)
        else:
            logger.info("  路由: SOCKS5h 127.0.0.1:%s", self.socks_port)
            socks_session = _build_session("socks", self.socks_port)
            try:
                socks = self._probe(
                    socks_session, "socks", "vpn", "Cloudflare-SOCKS", VPN_COMPARE_URL, cap
                )
            except requests.RequestException as e:
                socks = SpeedSample(
                    label="vpn/Cloudflare-SOCKS",
                    category="vpn",
                    route="socks",
                    url=VPN_COMPARE_URL,
                    ok=False,
                    error=str(e),
                )

        report.samples.extend([direct, socks])

        d_mbps = direct.download_mbps or 0.0
        s_mbps = socks.download_mbps or 0.0
        summary: dict[str, object] = {}

        if d_mbps > 0 and s_mbps > 0:
            ratio = round(s_mbps / d_mbps, 3)
            pct = round(ratio * 100, 1)
            summary = {
                "direct_mbps": d_mbps,
                "socks_mbps": s_mbps,
                "socks_over_direct_ratio": ratio,
            }
            if ratio >= 0.85:
                summary["conclusion"] = (
                    f"SOCKS 可达直连的 {pct:.0f}%，代理带宽正常，VPN 隧道有效。"
                )
            elif ratio >= 0.4:
                summary["conclusion"] = (
                    f"SOCKS 约为直连的 {pct:.0f}%，代理有损耗但可用，可检查节点负载或协议。"
                )
            else:
                summary["conclusion"] = (
                    f"SOCKS 仅约为直连的 {pct:.0f}%，代理瓶颈明显，建议换节点或检查分流。"
                )
        elif s_mbps > 0:
            summary = {
                "socks_mbps": s_mbps,
                "conclusion": "仅 SOCKS 测速成功；直连失败或路由未走同一出口。",
            }
        elif d_mbps > 0:
            if not self.socks_ready:
                summary = {
                    "direct_mbps": d_mbps,
                    "conclusion": f"直连 {d_mbps:.2f} Mbps；SOCKS 未测（{self.socks_hint}）。",
                }
            else:
                summary = {
                    "direct_mbps": d_mbps,
                    "conclusion": (
                        f"直连 {d_mbps:.2f} Mbps；SOCKS 失败（检查端口 {self.socks_port} 与代理是否开启）。"
                    ),
                }
        else:
            summary = {"conclusion": "VPN 对比两项均失败，请检查网络与代理。"}

        report.vpn_summary = summary
        report.socks_ready = self.socks_ready

    def run(self, skip_vpn: bool, skip_international: bool, domestic_route: str) -> RunReport:
        report = RunReport(
            timestamp=datetime.now().isoformat(timespec="seconds"),
            platform=platform.platform(),
            gateway=self.gateway,
            socks_ready=self.socks_ready,
        )
        logger.info("=" * 55)
        logger.info("31 模块：实时网速测试")
        if not self.socks_ready:
            logger.info("提示: %s（VPN SOCKS 对比将跳过或受限）", self.socks_hint)

        self.run_lan(report)
        self.run_category(report, "国内网", "domestic", DOMESTIC_PROBES, domestic_route)
        if not skip_international:
            self.run_category(report, "国外网", "international", INTERNATIONAL_PROBES, "direct")
        if not skip_vpn:
            self.run_vpn_compare(report)

        print_summary(report)
        logger.info("=" * 55)
        return report


def _resolve_target_ip(host: str):
    try:
        return ipaddress.ip_address(host)
    except ValueError:
        pass
    try:
        return ipaddress.ip_address(socket.gethostbyname(host))
    except OSError:
        return None


def _local_ip() -> str:
    try:
        return socket.gethostbyname(socket.gethostname())
    except OSError:
        return ""


def _lan_subnet_hint(gateway: str, local_ip: str) -> str:
    nets = _infer_scan_networks(gateway, local_ip)
    if not nets:
        return "—"
    return "、".join(str(n) for n in nets)


def _infer_scan_networks(gateway: str, local_ip: str) -> list[ipaddress.IPv4Network]:
    """根据网关与本机 IP 推断要扫描的私网网段（可跨多个 /24）。"""
    nets: list[ipaddress.IPv4Network] = []
    seen: set[str] = set()

    def add_host(ip_str: str) -> None:
        if not ip_str:
            return
        try:
            ip = ipaddress.ip_address(ip_str)
        except ValueError:
            return
        if not (ip.is_private or ip.is_loopback):
            return
        if ip.version != 4:
            return
        net = ipaddress.ip_network(f"{ip}/24", strict=False)
        key = str(net)
        if key not in seen:
            seen.add(key)
            nets.append(net)

    add_host(gateway)
    add_host(local_ip)
    return nets


@dataclass
class LanDevice:
    ip: str
    mac: str = ""
    ping_ms: float | None = None
    hostname: str = ""
    role: str = ""
    open_ports: list[int] = field(default_factory=list)
    notes: str = ""

    @property
    def display_name(self) -> str:
        if self.hostname:
            return self.hostname
        if self.role:
            return self.role
        return "未知设备"


def _read_arp_table() -> dict[str, str]:
    table: dict[str, str] = {}
    try:
        out = _check_output_hidden(["arp", "-a"], timeout=8)
        for line in out.splitlines():
            m = re.match(r"\s*(\d+\.\d+\.\d+\.\d+)\s+([0-9a-f-]+)\s+", line, re.I)
            if m:
                table[m.group(1)] = m.group(2).replace("-", ":").lower()
    except (subprocess.SubprocessError, OSError) as e:
        logger.debug("arp -a 失败: %s", e)
    return table


def _filter_arp_scope(
    arp: dict[str, str], networks: list[ipaddress.IPv4Network]
) -> dict[str, str]:
    """去掉组播/无效 MAC/网段外记录，避免对数百虚假项做探测。"""
    bad_mac = {"00:00:00:00:00:00", "ff:ff:ff:ff:ff:ff"}
    out: dict[str, str] = {}
    for ip, mac in arp.items():
        mac = (mac or "").lower()
        if not mac or mac in bad_mac or mac.startswith("01:00:5e"):
            continue
        try:
            addr = ipaddress.ip_address(ip)
        except ValueError:
            continue
        if not addr.is_private or addr.is_multicast or addr.is_loopback:
            continue
        if not any(addr in net for net in networks):
            continue
        out[ip] = mac
    return out


def _ping_once(ip: str) -> float | None:
    if platform.system() == "Windows":
        cmd = ["ping", "-n", "1", "-w", str(PING_TIMEOUT_MS_WIN), ip]
    else:
        cmd = ["ping", "-c", "1", "-W", "1", ip]
    try:
        out = _run_hidden(cmd, timeout=3)
        if out.returncode != 0:
            return None
        for pat in (
            r"(?:time|时间)\s*[=<]\s*(\d+(?:\.\d+)?)\s*ms",
            r"(?:time|时间)\s*[=<](\d+(?:\.\d+)?)\s*ms",
        ):
            m = re.search(pat, out.stdout, re.I)
            if m:
                return float(m.group(1))
        return 999.0
    except (subprocess.SubprocessError, OSError):
        return None


def _resolve_hostname_quick(ip: str) -> str:
    """仅 DNS 反向解析，不调用 nbtstat（避免弹窗且很慢）。"""
    old = socket.getdefaulttimeout()
    try:
        socket.setdefaulttimeout(0.6)
        name, _, _ = socket.gethostbyaddr(ip)
        if name and name != ip:
            return name.split(".")[0]
    except OSError:
        pass
    finally:
        socket.setdefaulttimeout(old)
    return ""


def _scan_ports(ip: str, ports: tuple[int, ...] = PORTS_QUICK) -> list[int]:
    return [p for p in ports if _tcp_port_open(ip, p)]


def _batch_port_scan(ips: list[str]) -> dict[str, list[int]]:
    """并行端口探测，避免对每台设备串行等待超时。"""
    if not ips:
        return {}
    result: dict[str, list[int]] = {ip: [] for ip in ips}
    workers = min(16, max(4, len(ips)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        fut_map = {pool.submit(_scan_ports, ip): ip for ip in ips}
        for fut in as_completed(fut_map):
            ip = fut_map[fut]
            try:
                result[ip] = fut.result()
            except Exception:
                result[ip] = []
    return result


def _batch_resolve_hostnames(ips: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    if not ips:
        return out
    with ThreadPoolExecutor(max_workers=min(8, len(ips))) as pool:
        fut_map = {pool.submit(_resolve_hostname_quick, ip): ip for ip in ips}
        for fut in as_completed(fut_map):
            ip = fut_map[fut]
            try:
                name = fut.result()
                if name:
                    out[ip] = name
            except Exception:
                pass
    return out


def _tcp_port_open(ip: str, port: int) -> bool:
    try:
        with socket.create_connection((ip, port), timeout=PORT_SCAN_TIMEOUT):
            return True
    except OSError:
        return False


def _guess_device_role(ip: str, gateway: str, local_ip: str, ports: list[int]) -> str:
    if gateway and ip == gateway:
        return "网关/路由器"
    if local_ip and ip == local_ip:
        return "本机"
    if 445 in ports or 139 in ports:
        return "Windows/NAS 文件共享"
    if 548 in ports or 5000 in ports:
        return "NAS/存储"
    if 80 in ports or 443 in ports or 8080 in ports:
        return "带 Web 服务设备"
    if 22 in ports:
        return "Linux/SSH 设备"
    return "局域网终端"


def _hosts_in_network(net: ipaddress.IPv4Network) -> list[str]:
    return [str(h) for h in net.hosts()]


def scan_lan_devices(gateway: str, local_ip: str, deep_scan: bool = False) -> list[LanDevice]:
    """
    快速模式（默认）：只 Ping「ARP 表里已有 + 网关/本机」，通常十秒内完成。
    深度模式：对全部 /24 地址 Ping（慢，508+ 地址，仅必要时使用）。
    """
    networks = _infer_scan_networks(gateway, local_ip)
    if not networks:
        logger.warning("未能推断扫描网段。")
        return []

    print("\n正在读取 ARP 邻居表…")
    arp_raw = _read_arp_table()
    arp = _filter_arp_scope(arp_raw, networks)

    ping_targets: set[str] = set(arp.keys())
    if gateway:
        ping_targets.add(gateway)
    if local_ip:
        ping_targets.add(local_ip)

    if deep_scan:
        for net in networks:
            ping_targets.update(_hosts_in_network(net))
        print(f"  深度扫描：Ping {len(ping_targets)} 个地址（约 1–3 分钟）…", flush=True)
    elif len(ping_targets) > FAST_PING_CAP:
        must = {x for x in (gateway, local_ip) if x}
        rest = sorted(ping_targets - must)[: max(0, FAST_PING_CAP - len(must))]
        ping_targets = must | set(rest)
        print(
            f"  快速扫描：ARP 邻居 {len(arp)} 台，仅 Ping 其中 {len(ping_targets)} 台"
            f"（其余仍列出 IP/MAC，不测延迟）…",
            flush=True,
        )
    else:
        print(
            f"  快速扫描：Ping {len(ping_targets)} 个邻居"
            f"（ARP {len(arp)} 条）…",
            flush=True,
        )

    responsive: dict[str, float] = {}
    targets = sorted(ping_targets)
    with ThreadPoolExecutor(max_workers=SCAN_PING_WORKERS) as pool:
        future_map = {pool.submit(_ping_once, ip): ip for ip in targets}
        done = 0
        for fut in as_completed(future_map):
            ip = future_map[fut]
            done += 1
            if deep_scan and done % 64 == 0:
                print(f"  进度 {done}/{len(targets)} …")
            try:
                ms = fut.result()
            except Exception:
                ms = None
            if ms is not None:
                responsive[ip] = ms

    candidate_ips = sorted(set(arp.keys()) | set(responsive.keys()))
    devices: list[LanDevice] = []

    print(f"  Ping 通 {len(responsive)} 台，待列出设备 {len(candidate_ips)} 台。", flush=True)

    # 快速模式：只对网关做端口探测（管理页）；深度模式：最多 15 台并行探测
    if deep_scan:
        port_ips = []
        if gateway:
            port_ips.append(gateway)
        for ip in sorted(responsive.keys()):
            if ip != gateway and len(port_ips) < 16:
                port_ips.append(ip)
        print(f"  端口探测（并行，最多 {len(port_ips)} 台）…", flush=True)
    else:
        port_ips = [gateway] if gateway else []
        print("  端口探测（仅网关，用于管理页地址）…", flush=True)

    port_map = _batch_port_scan(port_ips)

    name_ips = [x for x in (gateway, local_ip) if x]
    if deep_scan:
        extra = [ip for ip in sorted(responsive.keys()) if ip not in name_ips][:HOSTNAME_RESOLVE_LIMIT]
        name_ips.extend(extra)
    print(f"  解析主机名（{len(name_ips)} 台）…", flush=True)
    host_map = _batch_resolve_hostnames(name_ips)

    print("  正在生成设备列表…", flush=True)
    for idx, ip in enumerate(candidate_ips):
        if idx > 0 and idx % 40 == 0:
            print(f"    已处理 {idx}/{len(candidate_ips)} …", flush=True)

        ping_ms = responsive.get(ip)
        ports = port_map.get(ip, [])
        hostname = host_map.get(ip, "")
        role = _guess_device_role(ip, gateway, local_ip, ports)
        if not ports and ping_ms is not None and ip not in (gateway, local_ip):
            role = "在线终端"

        dev = LanDevice(
            ip=ip,
            mac=arp.get(ip, ""),
            ping_ms=round(ping_ms, 2) if ping_ms is not None and ping_ms < 900 else None,
            hostname=hostname,
            role=role,
            open_ports=ports,
        )
        if ip == gateway:
            dev.notes = "默认网关"
        elif ip == local_ip:
            dev.notes = "当前电脑"
        elif ping_ms is None and ip in arp:
            dev.notes = "ARP 可见（可能禁 Ping）"
        devices.append(dev)

    print(f"  完成，共 {len(devices)} 条记录。", flush=True)

    def sort_key(d: LanDevice) -> tuple[int, str]:
        if d.ip == gateway:
            return (0, d.ip)
        if d.ip == local_ip:
            return (1, d.ip)
        return (2, d.ip)

    devices.sort(key=sort_key)
    return devices


def _local_connection_stats() -> list[tuple[str, int]]:
    """本机当前外联 IP 连接数（无法直接代表「别人占带宽」，仅供本机排查）。"""
    counts: dict[str, int] = {}
    try:
        if platform.system() == "Windows":
            out = _check_output_hidden(["netstat", "-n"], timeout=10)
            for line in out.splitlines():
                if "ESTABLISHED" not in line.upper():
                    continue
                parts = line.split()
                if len(parts) < 3:
                    continue
                remote = parts[2]
                if remote.count(":") == 1:
                    host, _port = remote.rsplit(":", 1)
                else:
                    host = remote
                if host in ("0.0.0.0", "127.0.0.1", "::1", "*"):
                    continue
                counts[host] = counts.get(host, 0) + 1
        else:
            out = subprocess.check_output(["netstat", "-n"], text=True, errors="replace", timeout=10)
            for line in out.splitlines():
                if "ESTABLISHED" not in line:
                    continue
                parts = line.split()
                if len(parts) >= 5:
                    remote = parts[4]
                    host = remote.rsplit(":", 1)[0] if ":" in remote else remote
                    counts[host] = counts.get(host, 0) + 1
    except (subprocess.SubprocessError, OSError) as e:
        logger.debug("netstat 失败: %s", e)
    return sorted(counts.items(), key=lambda x: -x[1])[:15]


def _probe_router_admin_ports(gateway: str) -> list[int]:
    if not gateway:
        return []
    open_ports: list[int] = []
    for p in (443, 80, 8080, 8443):
        if _tcp_port_open(gateway, p):
            open_ports.append(p)
    return open_ports


def _router_admin_url(gateway: str, open_ports: list[int]) -> str:
    if not gateway or not open_ports:
        return f"http://{gateway}/" if gateway else ""
    if 443 in open_ports:
        return f"https://{gateway}/"
    if 8443 in open_ports:
        return f"https://{gateway}:8443/"
    if 8080 in open_ports:
        return f"http://{gateway}:8080/"
    return f"http://{gateway}/"


def run_lan_occupancy_survey(save_json: bool = False, deep_scan: bool = False) -> dict[str, object]:
    """
    排查「局域网里有哪些设备在线」及本机连接情况，并给出按 IP/MAC 限速的操作指引。
    说明：单台电脑无法直接测量「其他设备占了多少 Mbps」，需结合路由器流量统计。
    """
    gateway = _default_gateway()
    local_ip = _local_ip()
    hostname = socket.gethostname()
    devices = scan_lan_devices(gateway, local_ip, deep_scan=deep_scan)
    router_ports = _probe_router_admin_ports(gateway)
    admin_url = _router_admin_url(gateway, router_ports)
    conn_stats = _local_connection_stats()

    others = [
        d
        for d in devices
        if d.ip not in (gateway, local_ip) and (d.ping_ms is not None or d.mac)
    ]

    w = 60
    print()
    print("=" * w)
    print("  局域网占用排查报告".center(w))
    print("=" * w)
    print(f"  本机        {hostname}  ({local_ip or '—'})")
    print(f"  默认网关    {gateway or '—'}")
    print(f"  扫描网段    {_lan_subnet_hint(gateway, local_ip)}")
    if router_ports:
        print(f"  路由器管理  {admin_url}  （端口 {router_ports} 已开）")
    else:
        print("  路由器管理  未检测到 80/443（Web 管理可能关闭或使用非标准端口）")
    print("-" * w)
    print("  【在线设备】（Ping/ARP 扫描，约 1–2 分钟）")
    print(f"  {'IP':<16} {'MAC':<18} {'延迟':<8} {'类型/名称':<22} 备注")
    for d in devices:
        ping = f"{d.ping_ms} ms" if d.ping_ms is not None else "—"
        mac = (d.mac[:17] + "…") if len(d.mac) > 18 else (d.mac or "—")
        name = d.display_name[:20]
        note = d.notes or ""
        if d.open_ports:
            note = (note + " 端口" + ",".join(map(str, d.open_ports))).strip()
        flag = "  ← 重点" if d.ip in {x.ip for x in others} and d.ip != gateway else ""
        print(f"  {d.ip:<16} {mac:<18} {ping:<8} {name:<22} {note}{flag}")

    print("-" * w)
    print(f"  除网关/本机外，另有 {len(others)} 台设备在线。")
    if others:
        print("  若网速变慢，可在路由器「终端管理/流量统计」中对照下列 IP 查看实时流量：")
        for d in others[:12]:
            mac_hint = f"，MAC {d.mac}" if d.mac else ""
            print(f"    · {d.ip}  ({d.display_name}){mac_hint}")
    else:
        print("  未发现其他在线终端（或均禁 Ping）；慢网更可能出在外网/本机软件。")

    print("-" * w)
    print("  【本机外联连接数 Top】（只反映你这台电脑在连谁，不代表别人占带宽）")
    if conn_stats:
        for remote, cnt in conn_stats[:8]:
            tag = " [内网]" if _resolve_target_ip(remote) and ipaddress.ip_address(remote).is_private else ""
            print(f"    {remote:<28} {cnt:>4} 条连接{tag}")
    else:
        print("    （无 ESTABLISHED 连接或未读取到 netstat）")

    print("-" * w)
    print("  【重要说明】")
    print("    单机脚本不能测出「某 IP 当前下载多少 Mbps」。要抓占用大户，请：")
    print(f"    1) 浏览器打开路由器管理页: {admin_url or '(查看网关背面贴纸)'}")
    print("    2) 进入「终端管理 / 设备列表 / 流量统计 / QoS」查看各 IP 实时速率")
    print("    3) 对占用高的设备：按 IP 或 MAC 设置「下载限速」（不必断网）")
    print()
    print("  【为何不用 ARP 类工具在电脑上限速别人】")
    print("    · 常见「ARP 工具」多为欺骗/断网/中间人，不能稳定做到「1MB/s 降到 500KB/s」")
    print("    · 易误伤整网、引发掉线，且可能涉及未授权干预他人通信（有法律与合规风险）")
    print("    · 公平限速应在路由器 QoS 完成：设「下载上限」即可，对方仍可上网只是变慢")
    print("    · 例：对方约 1MB/s(≈8Mbps) → 在路由器将该设备下载上限设为 4Mbps(≈500KB/s)")
    print()
    print("  【针对性限速（通用步骤）】")
    print("    · 记录目标 IP 与 MAC（见上表）")
    print("    · 登录路由器 → QoS / 智能限速 / 家长控制 → 添加规则")
    print("    · 限制类型选「下载带宽上限」，例如 2–5 Mbps（按需要调整）")
    print("    · 保存后观察 5–10 分钟；勿对不属于自己的设备越权操作")
    print("=" * w)
    print()

    payload: dict[str, object] = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "gateway": gateway,
        "local_ip": local_ip,
        "admin_url": admin_url,
        "devices": [asdict(d) for d in devices],
        "other_device_count": len(others),
        "local_connections": [{"remote": r, "count": c} for r, c in conn_stats],
    }

    if save_json:
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = DEFAULT_OUTPUT_DIR / f"lan_survey_{ts}.json"
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("报告已保存: %s", out)

    return payload


def _prompt_int(prompt: str, default: int, min_v: int, max_v: int) -> int:
    raw = input(f"{prompt} [{default}]: ").strip()
    if not raw:
        return default
    try:
        val = int(raw)
        return max(min_v, min(val, max_v))
    except ValueError:
        return default


def run_interactive() -> None:
    """无命令行参数时：运行后交互选择功能。"""
    print()
    print("=" * 50)
    print("  31 网速工具 — 请选择功能")
    print("=" * 50)
    print("  1. 完整测速（局域网 + 国内 + 国外 + VPN）")
    print("  2. 快速测速（跳过 VPN 对比）")
    print("  3. 仅局域网（网关 Ping + 可选内网 URL 单次测速）")
    print("  4. 局域网占用排查（扫在线设备 IP/MAC + 限速指引）【推荐】")
    print("  0. 退出")
    print("-" * 50)
    print("说明：网速慢、查「谁占网」请用菜单 4；")
    print("      在路由器管理页看各设备流量，并按 IP/MAC 设下载上限。")
    print()

    choice = input("请输入序号 [1]: ").strip() or "1"
    if choice == "0":
        return

    if choice == "4":
        deep = input("深度扫描整段 /24？(y/N，默认 N 快速模式): ").strip().lower() == "y"
        save = input("是否保存 JSON 报告？(y/N): ").strip().lower() == "y"
        run_lan_occupancy_survey(save_json=save, deep_scan=deep)
        return

    socks_port = _prompt_int("SOCKS 端口", 10808, 1, 65535)
    timeout_s = _prompt_int("请求超时(秒)", 15, 5, 60)

    lan_url = ""
    skip_vpn = False
    skip_intl = False

    if choice == "2":
        skip_vpn = True
    elif choice == "3":
        skip_vpn = True
        skip_intl = True
        lan_url = input("内网测速 URL（可回车跳过）: ").strip()
    elif choice != "1":
        print("无效序号，按完整测速执行。")

    if choice in ("1", "2") and not lan_url:
        lan_url = input("内网测速 URL（可回车跳过）: ").strip()

    save = input("是否保存 JSON 报告？(y/N): ").strip().lower() == "y"
    quiet = input("仅显示最终汇总？(y/N): ").strip().lower() == "y"
    if quiet:
        logger.setLevel(logging.WARNING)

    tester = NetworkSpeedTester(
        socks_port=socks_port,
        timeout_s=timeout_s,
        max_bytes=10_485_760,
        live=not quiet,
        lan_url=lan_url,
    )
    report = tester.run(
        skip_vpn=skip_vpn,
        skip_international=skip_intl,
        domestic_route="direct",
    )
    if save:
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = DEFAULT_OUTPUT_DIR / f"speed_test_{ts}.json"
        out.write_text(
            json.dumps(
                {
                    **{k: v for k, v in asdict(report).items() if k != "samples"},
                    "samples": [asdict(s) for s in report.samples],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        logger.info("报告已保存: %s", out)


def main_cli() -> None:
    parser = argparse.ArgumentParser(description="31 模块：实时网速（局域网 / 国内 / 国外 / VPN）")
    parser.add_argument("--socks-port", type=int, default=10808, help="SOCKS5 端口（VPN 对比）")
    parser.add_argument("--timeout", type=int, default=15, help="单次请求超时（秒）")
    parser.add_argument("--max-mb", type=float, default=10.0, help="单探针最大下载量（MB）")
    parser.add_argument(
        "--domestic-route",
        choices=["direct", "socks"],
        default="direct",
        help="国内网探针使用的路由（默认直连）",
    )
    parser.add_argument("--lan-url", default="", help="可选：内网测速文件 URL（如 NAS/路由器）")
    parser.add_argument("--skip-vpn", action="store_true", help="跳过 VPN 直连 vs SOCKS 对比")
    parser.add_argument("--skip-international", action="store_true", help="跳过国外网探针")
    parser.add_argument("--no-live", action="store_true", help="关闭下载过程中的实时 Mbps 刷新")
    parser.add_argument("--save-json", action="store_true", help="保存 JSON 到 output/")
    parser.add_argument("--quiet", action="store_true", help="仅输出最终汇总（隐藏过程日志）")
    args = parser.parse_args()

    if args.quiet:
        logger.setLevel(logging.WARNING)

    tester = NetworkSpeedTester(
        socks_port=args.socks_port,
        timeout_s=args.timeout,
        max_bytes=int(args.max_mb * 1024 * 1024),
        live=not args.no_live,
        lan_url=args.lan_url,
    )
    report = tester.run(
        skip_vpn=args.skip_vpn,
        skip_international=args.skip_international,
        domestic_route=args.domestic_route,
    )

    if args.save_json:
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = DEFAULT_OUTPUT_DIR / f"speed_test_{ts}.json"
        payload = {
            **{k: v for k, v in asdict(report).items() if k != "samples"},
            "samples": [asdict(s) for s in report.samples],
        }
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("报告已保存: %s", out)

    if not any(s.ok for s in report.samples):
        sys.exit(1)


def main() -> None:
    if len(sys.argv) <= 1:
        run_interactive()
    else:
        main_cli()


if __name__ == "__main__":
    main()
