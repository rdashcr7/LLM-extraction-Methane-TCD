"""Shared AWS Bedrock Converse transport for every publisher extractor."""

import os
import random
import time
from typing import Any, Callable


def selected_provider() -> str:
    provider = os.environ.get("LLM_PROVIDER", "openai").strip().lower()
    if provider not in {"openai", "bedrock"}:
        raise ValueError("LLM_PROVIDER must be 'openai' or 'bedrock'.")
    return provider


def _resolve_model_id(model_id: str, region: str) -> str:
    """Use the regional inference-profile prefix when one is not already supplied."""
    raw = (model_id or "").strip()
    if not raw or raw.startswith(("us.", "eu.", "ap.", "au.", "apac.", "global.")):
        return raw
    prefixes = {
        "us-east-1": "us", "us-east-2": "us", "us-west-1": "us", "us-west-2": "us",
        "eu-west-1": "eu", "eu-west-2": "eu", "eu-west-3": "eu", "eu-central-1": "eu", "eu-central-2": "eu",
        "ap-northeast-1": "ap", "ap-southeast-1": "ap", "ap-southeast-2": "ap", "ap-south-1": "ap", "ap-south-2": "ap",
    }
    return f"{prefixes.get(region, 'us')}.{raw}"


def _client_and_model() -> tuple[Any, str]:
    region = os.environ.get("AWS_REGION", "").strip()
    model_id = os.environ.get("BEDROCK_MODEL_ID", os.environ.get("MODEL_ID", "")).strip()
    if not region or not model_id:
        raise RuntimeError("Bedrock requires AWS_REGION and BEDROCK_MODEL_ID (or MODEL_ID).")
    try:
        import boto3
    except ImportError as exc:
        raise RuntimeError("AWS Bedrock requires boto3. Install it with: pip install boto3") from exc

    kwargs: dict[str, Any] = {"service_name": "bedrock-runtime", "region_name": region}
    access_key = os.environ.get("AWS_ACCESS_KEY_ID", "").strip()
    secret_key = os.environ.get("AWS_SECRET_ACCESS_KEY", "").strip()
    session_token = os.environ.get("AWS_SESSION_TOKEN", "").strip()
    if access_key and secret_key:
        kwargs["aws_access_key_id"] = access_key
        kwargs["aws_secret_access_key"] = secret_key
    if session_token:
        kwargs["aws_session_token"] = session_token
    return boto3.client(**kwargs), _resolve_model_id(model_id, region)


def _output_text(response: dict[str, Any]) -> str:
    parts = [item.get("text", "") for item in response.get("output", {}).get("message", {}).get("content", []) if isinstance(item, dict)]
    text = "\n".join(str(part) for part in parts if str(part).strip()).strip()
    if not text:
        raise RuntimeError("Bedrock response contained no output text.")
    return text


def converse(
    *,
    messages: list[dict[str, Any]],
    system_prompt: str,
    max_tokens: int,
    temperature: float,
    call_delay_seconds: float,
    max_retries: int,
    base_delay_seconds: float,
    max_delay_seconds: float,
    jitter_seconds: float,
    is_throttling: Callable[[BaseException], bool],
    record_usage: Callable[[dict[str, Any]], None],
    raise_rate_limit: Callable[[str], None],
) -> str:
    """Call any supported Bedrock Converse model with the extractor's native messages."""
    client, model_id = _client_and_model()
    last_error: BaseException | None = None
    for attempt in range(1, max(1, max_retries + 1) + 1):
        if call_delay_seconds > 0:
            time.sleep(call_delay_seconds)
        try:
            response = client.converse(
                modelId=model_id,
                system=[{"text": system_prompt}],
                messages=messages,
                inferenceConfig={"maxTokens": max_tokens, "temperature": temperature},
            )
            record_usage(response)
            stop_reason = str(response.get("stopReason", "")).lower()
            if stop_reason in {"max_tokens", "length", "model_context_window_exceeded"}:
                print(f"      [warning] Bedrock response stopped because of {stop_reason}; parser salvage will be used if possible")
            return _output_text(response)
        except Exception as exc:
            if not is_throttling(exc):
                raise
            last_error = exc
            if attempt >= max(1, max_retries + 1):
                break
            delay = min(base_delay_seconds * (2 ** (attempt - 1)), max_delay_seconds)
            if jitter_seconds > 0:
                delay += random.uniform(0.0, jitter_seconds)
            print(f"      [throttle] Bedrock rate/quota/throughput limit; sleeping {delay:.1f}s before application retry {attempt}/{max_retries}")
            time.sleep(max(0.0, delay))
    raise_rate_limit(str(last_error) if last_error else "unknown Bedrock throttling/rate/quota error")
    raise AssertionError("raise_rate_limit must raise an exception")
