import json
import logging
import re
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from typing import Optional, Dict, Any

from settings import load_settings, check_environment

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)


class ClinicalDataGuard:
    """
    临床严重不良事件 (SAE) 结构化提取引擎。
    对接 OpenAI 兼容 Chat Completions 端点，输出确定性 JSON 字段。
    """

    def __init__(self, base_url: str, token: str, model_id: str):
        if not base_url or not isinstance(base_url, str):
            raise ValueError("base_url 不能为空且必须为字符串。")
        if not token or not isinstance(token, str):
            raise ValueError("token 不能为空且必须为字符串。")
        if not model_id or not isinstance(model_id, str):
            raise ValueError("model_id 不能为空且必须为字符串。")

        self.endpoint = f"{base_url.rstrip('/')}/v1/chat/completions"
        self.model_id = model_id

        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}"
        })

        retry_strategy = Retry(
            total=3,
            backoff_factor=1.0,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["POST"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

        self.system_prompt = """
        你是一名严谨的临床医学数据审查专家。请从输入的非结构化临床病历或随访记录中，提取严重不良事件（SAE）的关键要素。
        必须严格且仅输出以下 JSON 结构，若文本中未提供对应信息，对应字段的值必须为 null。不要包含任何附加解释。
        {
            "sae_term": "严重不良事件标准名称",
            "onset_date": "发生日期 (格式: YYYY-MM-DD)",
            "resolution_date": "转归日期 (格式: YYYY-MM-DD)",
            "severity_grade": "CTCAE 严重程度分级 (如: 3级, 4级)",
            "causality": "与试验药物的关联性 (肯定有关/可能有关/可能无关/肯定无关/无法判定)",
            "action_taken": "对试验药物采取的措施 (如: 停药/减量/维持不变)",
            "outcome": "事件转归结果 (恢复/未恢复/死亡/未知)"
        }
        """

    @staticmethod
    def _extract_json_text(raw_content: str) -> Optional[str]:
        if not raw_content:
            return None

        content = raw_content.strip()
        if content.startswith("```json"):
            content = content[7:].strip()
        if content.startswith("```"):
            content = content[3:].strip()
        if content.endswith("```"):
            content = content[:-3].strip()

        try:
            json.loads(content)
            return content
        except json.JSONDecodeError as e:
            # raw_content 不是合法 JSON，将回退到 regex 提取首个 {...}。
            # 保留 debug 日志以便审计模型返回结构是否符合预期。
            logging.debug("raw_content 不是合法 JSON，回退 regex: %s", e)

        match = re.search(r"\{.*\}", content, flags=re.DOTALL)
        if match:
            return match.group(0).strip()
        return None

    @staticmethod
    def _normalize_result(data: Dict[str, Any]) -> Dict[str, Any]:
        required_fields = [
            "sae_term",
            "onset_date",
            "resolution_date",
            "severity_grade",
            "causality",
            "action_taken",
            "outcome",
        ]
        normalized = {k: data.get(k) for k in required_fields}
        for k, v in data.items():
            if k not in normalized:
                normalized[k] = v
        return normalized

    def extract(self, clinical_text: str, timeout: int = 120) -> Optional[Dict[str, Any]]:
        if not isinstance(clinical_text, str) or not clinical_text.strip():
            logging.warning("输入文本为空或非字符串，跳过提取。")
            return None

        payload = {
            "model": self.model_id,
            "messages": [
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": f"病历文本：\n{clinical_text}"}
            ],
            "temperature": 0.0,
            "top_p": 0.1,
            "response_format": {"type": "json_object"}
        }

        logging.info(f"发起 SAE 实体提取请求 | 文本长度: {len(clinical_text)} 字符")

        try:
            response = self.session.post(
                self.endpoint,
                json=payload,
                timeout=timeout
            )
            response.raise_for_status()

            response_json = response.json()
            choices = response_json.get("choices") or []
            if not choices:
                logging.error("模型返回中缺少 choices 字段。")
                return None

            message = choices[0].get("message") if isinstance(choices[0], dict) else None
            raw_content = (message or {}).get("content")
            json_text = self._extract_json_text(raw_content if isinstance(raw_content, str) else "")
            if not json_text:
                logging.error("模型返回内容中未识别到有效 JSON。")
                logging.debug("原始输出: %s", raw_content)
                return None

            structured_data = json.loads(json_text)
            if not isinstance(structured_data, dict):
                logging.error("模型返回 JSON 不是对象结构。")
                return None

            structured_data = self._normalize_result(structured_data)
            logging.info("结构化提取成功。")
            return structured_data

        except requests.exceptions.RequestException as e:
            logging.error(f"网关通信失败或发生异常重试耗尽: {e}")
            return None
        except (json.JSONDecodeError, TypeError, ValueError) as e:
            logging.error(f"模型输出格式异常或非标准 JSON: {e}")
            logging.debug(f"原始错误输出: {raw_content if 'raw_content' in locals() else '无输出'}")
            return None


if __name__ == "__main__":
    settings = load_settings()
    env = check_environment(settings)
    logging.info(f"环境自检摘要 | Tesseract: {env['tesseract']['ok']} | Poppler: {env['poppler']['ok']} | Token已配置: {env['api']['token_set']}")

    if not settings.api_token:
        raise SystemExit(
            "未检测到 SAE_API_TOKEN。\n"
            "请先在环境变量中设置，例如（PowerShell）：\n"
            "  $env:SAE_API_TOKEN='你的Token'\n"
            "然后重试。"
        )

    extractor = ClinicalDataGuard(
        base_url=settings.api_base_url,
        token=settings.api_token,
        model_id=settings.model_id
    )

    test_clinical_text = """
    受试者编号 S-0012，于2026年3月10日因“持续高热伴胸闷”入院。
    临床诊断为心肌炎（CTCAE 3级）。研究者评估该不良事件与试验疫苗可能有关。
    予以暂停第二剂疫苗接种，并给予营养心肌及对症支持治疗。
    随访至2026年3月16日，受试者症状完全消退，复查心肌酶谱恢复正常，事件评估为已恢复。
    """

    result = extractor.extract(test_clinical_text)

    if result:
        print("\n=== SAE 结构化输出结果 ===")
        print(json.dumps(result, indent=4, ensure_ascii=False))
    else:
        print("\n=== 提取失败，请检查日志与网络 ===")
