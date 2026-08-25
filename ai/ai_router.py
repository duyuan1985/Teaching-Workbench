"""AI 路由：本地 Ollama + 在线模型（智谱 GLM / OpenRouter）自动 fallback。"""

import json
import os
import time
import urllib.error
import urllib.request
from typing import Optional, Dict, List
from dataclasses import dataclass

import store


@dataclass
class AIConfig:
    name: str
    type: str
    url: str
    model: str
    api_key: Optional[str] = None
    timeout: int = 60
    enabled: bool = True
    cost_per_mtok: float = 0.0


def current_url():
    return os.environ.get("WORKBENCH_OLLAMA_URL") or store.get_setting("ollama_url", "http://127.0.0.1:11434")


def current_model():
    return os.environ.get("WORKBENCH_OLLAMA_MODEL") or store.get_setting("ollama_model", "qwen3:8b")


def current_preference():
    """AI源优先级：online（默认，智谱GLM优先）/ local（Ollama优先）/ auto（按复杂度）。"""
    return os.environ.get("WORKBENCH_AI_PREFERENCE") or store.get_setting("ai_model_preference", "online")


def installed_models(timeout=2.0):
    try:
        request = urllib.request.Request(f"{current_url().rstrip('/')}/api/tags", method="GET")
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
        return [item.get("name", "") for item in payload.get("models", []) if item.get("name")]
    except (OSError, urllib.error.URLError, json.JSONDecodeError):
        return []


def ollama_available(timeout=1.5):
    try:
        request = urllib.request.Request(f"{current_url().rstrip('/')}/api/tags", method="GET")
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status == 200
    except (OSError, urllib.error.URLError):
        return False


def _build_sources():
    ollama_url = current_url().rstrip('/')
    ollama_model = current_model()
    return [
        AIConfig(
            name="本地7B",
            type="local",
            url=f"{ollama_url}/api/chat",
            model=ollama_model,
            timeout=200,
            cost_per_mtok=0.0,
        ),
        AIConfig(
            name="智谱GLM",
            type="online",
            url="https://open.bigmodel.cn/api/paas/v4/chat/completions",
            model="GLM-4-Flash-250414",   #GLM-Z1-Flash
            api_key=os.getenv("ZHIPU_API_KEY", ""),
            timeout=90,
            cost_per_mtok=0.0,
        ),
    ]


def _default_headers():
    """Cloudflare会按浏览器签名拦截Python默认UA（403 error 1010），需伪装浏览器UA。"""
    return {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
    }


def _post_json(url: str, payload: Dict, timeout: int, headers: Optional[Dict] = None):
    request_headers = _default_headers()
    request_headers.update(headers or {})
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=request_headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
            return response.status, json.loads(raw), raw
    except urllib.error.HTTPError as error:
        raw = error.read().decode("utf-8", errors="replace")
        return error.code, {}, raw


class AIRouter:
    def __init__(self):
        self.sources = []
        for s in _build_sources():
            if s.type == "local":
                self.sources.append(s)
            elif s.api_key and len(s.api_key) > 10:
                self.sources.append(s)

        self.stats = {s.name: {"calls": 0, "fails": 0, "total_time": 0} for s in self.sources}
        self._warmed_up = False
        # 本地源一旦超时/失败，本进程内后续调用直接跳过（避免每次都先耗数分钟等超时）；
        # 无在线源可用时仍保留本地，让它继续尝试
        self._local_disabled = False

    def _warmup_ollama(self, config: AIConfig):
        if self._warmed_up:
            return
        try:
            print("[预热] 正在加载本地模型到内存（约60-90秒，请等待）...")
            base_url = config.url.rsplit("/api/", 1)[0]
            _post_json(
                f"{base_url}/api/generate",
                {
                    "model": config.model,
                    "prompt": "hello",
                    "stream": False,
                    "options": {"num_predict": 1}
                },
                timeout=120,
            )
            self._warmed_up = True
            print("[预热] 模型已在内存中")
        except Exception as e:
            print(f"[预热] 失败: {e}")

    def _call_ollama(self, config: AIConfig, messages: List[Dict], system: str) -> Optional[str]:
        try:
            self._warmup_ollama(config)
            payload = {
                "model": config.model,
                "messages": [{"role": "system", "content": system}] + messages,
                "stream": False,
                "options": {"temperature": 0.4, "num_ctx": 8192}
            }
            status, response, _ = _post_json(config.url, payload, config.timeout)
            if status != 200:
                print(f"  [本地Ollama HTTP {status}]")
                return None
            return response["message"]["content"]
        except Exception as e:
            print(f"  [本地Ollama 错误] {e}")
            return None

    def _stream_chat(self, config: AIConfig, payload: Dict, headers: Dict) -> Optional[str]:
        """流式调用OpenAI兼容接口。分块持续到达即可不断重置读超时，
        慢速中转站不再因整包生成时间过长而超时。"""
        request_headers = _default_headers()
        request_headers.update(headers)
        request = urllib.request.Request(
            config.url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=request_headers,
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=config.timeout) as response:
                content_type = (response.headers.get("Content-Type") or "").lower()
                if "text/event-stream" not in content_type:
                    # 中转站忽略stream参数，直接返回了完整JSON
                    raw = response.read().decode("utf-8", errors="replace")
                    try:
                        return json.loads(raw)["choices"][0]["message"]["content"]
                    except (json.JSONDecodeError, KeyError, IndexError):
                        print(f"  [{config.name}] 响应异常: {raw[:100]}")
                        return None
                parts = []
                for raw_line in response:
                    line = raw_line.decode("utf-8", errors="replace").strip()
                    if not line.startswith("data:"):
                        continue
                    data_str = line[5:].strip()
                    if data_str == "[DONE]":
                        break
                    try:
                        chunk = json.loads(data_str)
                        delta = chunk["choices"][0].get("delta") or {}
                        parts.append(delta.get("content") or "")
                    except (json.JSONDecodeError, KeyError, IndexError, AttributeError):
                        continue
                content = "".join(parts)
                return content if content else None
        except urllib.error.HTTPError as error:
            raw = error.read().decode("utf-8", errors="replace")
            print(f"  [{config.name}] HTTP {error.code}: {raw[:100]}")
            return None
        except Exception as e:
            print(f"  [{config.name}] 流式错误: {e}")
            return None

    def _call_online(self, config: AIConfig, messages: List[Dict], system: str) -> Optional[str]:
        headers = {
            "Authorization": f"Bearer {config.api_key}",
        }
        payload = {
            "model": config.model,
            "messages": [{"role": "system", "content": system}] + messages,
            "temperature": 0.4,
            "max_tokens": 4096,
        }
        # 优先流式：非流式要等全部内容生成完才返回，慢速中转站极易整包读超时
        content = self._stream_chat(config, {**payload, "stream": True}, headers)
        if content:
            return content
        # 流式失败，回退非流式
        try:
            status, response, raw = _post_json(
                config.url,
                payload,
                timeout=config.timeout,
                headers=headers,
            )
            if status == 401:
                print(f"  [{config.name}] 401 Key错误")
                return None
            elif status == 402:
                print(f"  [{config.name}] 402 余额不足")
                return None
            elif status == 429:
                print(f"  [{config.name}] 429 限流")
                return None
            elif status != 200:
                print(f"  [{config.name}] HTTP {status}: {raw[:100]}")
                return None
            return response["choices"][0]["message"]["content"]
        except Exception as e:
            print(f"  [{config.name}] 错误: {e}")
            return None

    def ask(self, prompt: str, system: str = "你是一个专业的编程助手",
            prefer_local: bool = False, force_online: bool = False,
            show_details: bool = True) -> Dict:

        online_sources = [s for s in self.sources if s.type == "online"]
        local_sources = [s for s in self.sources if s.type == "local"]

        if force_online:
            candidates = online_sources
        elif prefer_local:
            candidates = local_sources + online_sources
        else:
            preference = current_preference()
            if preference == "local":
                candidates = local_sources + online_sources
            elif preference == "auto":
                complex_kw = ["架构", "设计", "方案", "整体", "完整项目", "数据库", "技术选型"]
                is_complex = any(kw in prompt for kw in complex_kw)
                candidates = (online_sources + local_sources) if is_complex else (local_sources + online_sources)
            else:
                # online（默认）：本机GPU跑不动7B模型，本地Ollama只在在线源失败时兜底
                candidates = online_sources + local_sources

        if not candidates:
            return {"success": False, "error": "没有可用的AI源", "content": ""}

        if self._local_disabled and any(s.type == "online" for s in candidates):
            candidates = [s for s in candidates if s.type != "local"]

        messages = [{"role": "user", "content": prompt}]

        for src in candidates:
            if show_details:
                cost = "免费" if src.cost_per_mtok == 0 else f"{src.cost_per_mtok}元/百万token"
                print(f"\n[尝试] {src.name} ({src.type}, {cost})")

            start = time.time()
            if src.type == "local":
                result = self._call_ollama(src, messages, system)
            else:
                result = self._call_online(src, messages, system)

            elapsed = time.time() - start
            self.stats[src.name]["calls"] += 1
            self.stats[src.name]["total_time"] += elapsed

            if result:
                if show_details:
                    print(f"[成功] {src.name} 耗时{elapsed:.1f}秒")
                return {"success": True, "source": src.name, "model": src.model, "content": result}
            else:
                self.stats[src.name]["fails"] += 1
                if src.type == "local":
                    self._local_disabled = True
                if show_details:
                    print(f"[失败] {src.name}，尝试下一个...")

        return {"success": False, "error": "所有源都失败了", "content": ""}

    def print_stats(self):
        print("\n" + "=" * 50)
        print("AI 使用统计")
        print("=" * 50)
        for name, stat in self.stats.items():
            avg = stat["total_time"] / stat["calls"] if stat["calls"] > 0 else 0
            print(f"{name:12s} | 调用{stat['calls']:3d}次 | 失败{stat['fails']:2d}次 | 平均{avg:.1f}秒")


_router = None


def _get_router():
    global _router
    if _router is None:
        _router = AIRouter()
    return _router


def ask_result(prompt: str, system: str = "你是一个专业的编程助手",
               prefer_local: bool = False, force_online: bool = False,
               show_details: bool = True) -> Dict:
    return _get_router().ask(prompt, system, prefer_local, force_online, show_details)


def ask(prompt: str, system: str = "你是一个专业的编程助手",
        prefer_local: bool = False, force_online: bool = False,
        show_details: bool = True) -> str:
    result = ask_result(prompt, system, prefer_local, force_online, show_details)
    return result["content"] if result["success"] else f"错误: {result.get('error')}"
