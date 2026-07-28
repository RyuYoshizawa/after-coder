"""
llm_client.py
プロバイダ非依存のLLM呼び出しモジュール。
tool_use / function_calling / フォールバック（JSON文字列）を自動切替。
"""

import json
import re
import time


# グローバルトークンカウンター
_token_usage = {'input': 0, 'output': 0}

def reset_token_usage():
    """トークンカウンターをリセット"""
    _token_usage['input']  = 0
    _token_usage['output'] = 0

def get_token_usage() -> dict:
    """現在のトークン使用量を返す"""
    return dict(_token_usage)


# 直近の呼び出し失敗理由（call_llmがNoneを返した際にUI側で原因表示するために使う）
_last_error = None

def get_last_error() -> str | None:
    """call_llmが直近でNoneを返した理由（例外メッセージ）を返す"""
    return _last_error

def calc_cost_jpy(input_tokens: int, output_tokens: int, model: str) -> float:
    """
    概算コストを円で返す（1USD=150円換算）
    claude-sonnet-4-6: input $3/1M, output $15/1M
    """
    rate = 150  # USD→JPY
    if 'sonnet' in model:
        cost = (input_tokens / 1_000_000 * 3 + output_tokens / 1_000_000 * 15) * rate
    elif 'haiku' in model:
        cost = (input_tokens / 1_000_000 * 0.25 + output_tokens / 1_000_000 * 1.25) * rate
    else:
        cost = (input_tokens / 1_000_000 * 3 + output_tokens / 1_000_000 * 15) * rate
    return round(cost, 1)


def call_llm(client, prompt: str, schema: dict, provider: str, model: str) -> dict | list | None:
    """
    構造化出力でLLMを呼び出す。
    tool_use / function_calling → 失敗時はJSONフォールバックの順で試みる。
    トークン使用量をグローバルカウンターに累積する。
    """
    global _last_error
    last_reason = None
    # 1回目は再現性のためtemperature=0固定。空結果で失敗した場合、
    # 同一プロンプト×temperature=0では毎回同じ結果になり得るため、
    # リトライ時は少しずつtemperatureを上げて出力にばらつきを持たせる
    temperatures = [0, 0.4, 0.7]

    for attempt in range(3):
        try:
            temperature = temperatures[attempt]
            if provider == 'Anthropic':
                result, usage = _call_anthropic(client, prompt, schema, model, temperature)
            elif provider == 'OpenAI':
                result, usage = _call_openai(client, prompt, schema, model, temperature)
            else:
                raise ValueError(f'未対応のプロバイダ: {provider}')

            if usage:
                _token_usage['input']  += usage.get('input', 0)
                _token_usage['output'] += usage.get('output', 0)

            if result:
                _last_error = None
                return result

            if result is None:
                last_reason = 'モデルが構造化データを返しませんでした（tool_use/JSON抽出とも失敗）'
            else:
                last_reason = f'モデルが空の結果を返しました（{result!r}）'

        except Exception as e:
            last_reason = f'{type(e).__name__}: {e}'
            print(f'\n  [リトライ{attempt+1}] {last_reason}')
            time.sleep(1)

    _last_error = last_reason
    print(f'\n  [警告] 構造化出力に失敗しました。このバッチをスキップします。理由: {last_reason}')
    return None


def _call_anthropic(client, prompt: str, schema: dict, model: str, temperature: float = 0) -> tuple:
    """Anthropic tool_use を使った構造化呼び出し"""
    tool = {
        'name': 'output_result',
        'description': 'プロンプトの指示に従って結果を構造化して返す',
        'input_schema': schema,
    }
    response = client.messages.create(
        model=model,
        max_tokens=8192,
        temperature=temperature,
        tools=[tool],
        tool_choice={'type': 'tool', 'name': 'output_result'},
        messages=[{'role': 'user', 'content': prompt}]
    )
    usage = {
        'input':  response.usage.input_tokens,
        'output': response.usage.output_tokens,
    }
    if response.stop_reason == 'max_tokens':
        raise RuntimeError(
            f'出力がmax_tokens上限で打ち切られました（入力{usage["input"]}トークン/出力{usage["output"]}トークン）。'
            'データ量やコード数上限を減らすか、分割して再実行してください。'
        )
    for block in response.content:
        if block.type == 'tool_use' and block.name == 'output_result':
            return block.input, usage
    for block in response.content:
        if block.type == 'text':
            result = _parse_json_text(block.text)
            if result is not None:
                return result, usage
    return None, usage


def _call_openai(client, prompt: str, schema: dict, model: str, temperature: float = 0) -> tuple:
    """OpenAI function_calling を使った構造化呼び出し"""
    tool = {
        'type': 'function',
        'function': {
            'name': 'output_result',
            'description': 'プロンプトの指示に従って結果を構造化して返す',
            'parameters': schema,
        }
    }
    response = client.chat.completions.create(
        model=model,
        max_tokens=8192,
        temperature=temperature,
        tools=[tool],
        tool_choice={'type': 'function', 'function': {'name': 'output_result'}},
        messages=[{'role': 'user', 'content': prompt}]
    )
    usage = {
        'input':  response.usage.prompt_tokens,
        'output': response.usage.completion_tokens,
    }
    if response.choices[0].finish_reason == 'length':
        raise RuntimeError(
            f'出力がmax_tokens上限で打ち切られました（入力{usage["input"]}トークン/出力{usage["output"]}トークン）。'
            'データ量やコード数上限を減らすか、分割して再実行してください。'
        )
    msg = response.choices[0].message
    if msg.tool_calls:
        for tc in msg.tool_calls:
            if tc.function.name == 'output_result':
                return json.loads(tc.function.arguments), usage
    if msg.content:
        return _parse_json_text(msg.content), usage
    return None, usage


def _parse_json_text(text: str) -> dict | list | None:
    """テキストからJSONを抽出するフォールバック処理"""
    text = text.strip()
    text = re.sub(r'^```json\s*', '', text)
    text = re.sub(r'^```\s*',     '', text)
    text = re.sub(r'\s*```$',     '', text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # 部分的に有効なJSONを救出
    depth_c=depth_s=last=0; in_s=esc=False
    for i, ch in enumerate(text):
        if esc: esc=False; continue
        if ch=='\\' and in_s: esc=True; continue
        if ch=='"': in_s=not in_s; continue
        if in_s: continue
        if ch=='{': depth_c+=1
        elif ch=='}':
            depth_c-=1
            if depth_c==0 and depth_s==0: last=i+1
        elif ch=='[': depth_s+=1
        elif ch==']':
            depth_s-=1
            if depth_c==0 and depth_s==0: last=i+1
    if last > 0:
        try:
            return json.loads(text[:last])
        except:
            pass
    return None


def make_client(provider: str, api_key: str):
    """プロバイダに応じたクライアントを生成する"""
    if provider == 'Anthropic':
        import anthropic
        return anthropic.Anthropic(api_key=api_key)
    elif provider == 'OpenAI':
        import openai
        return openai.OpenAI(api_key=api_key)
    else:
        raise ValueError(f'未対応のプロバイダ: {provider}')
