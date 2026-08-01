"""
llm_client.py
プロバイダ非依存のLLM呼び出しモジュール。
tool_use / function_calling / フォールバック（JSON文字列）を自動切替。
"""

import json
import re
import time


# グローバルトークンカウンター（キャッシュ読み込み/書き込み分と累積コストも保持）
_token_usage = {'input': 0, 'output': 0, 'cache_read': 0, 'cache_creation': 0, 'cost_jpy': 0.0}

def reset_token_usage():
    """トークンカウンターをリセット"""
    _token_usage['input']          = 0
    _token_usage['output']         = 0
    _token_usage['cache_read']     = 0
    _token_usage['cache_creation'] = 0
    _token_usage['cost_jpy']       = 0.0

def get_token_usage() -> dict:
    """現在のトークン使用量（キャッシュ内訳・累積コスト込み）を返す"""
    return dict(_token_usage)


# 直近の呼び出し失敗理由（call_llmがNoneを返した際にUI側で原因表示するために使う）
_last_error = None

def get_last_error() -> str | None:
    """call_llmが直近でNoneを返した理由（例外メッセージ）を返す"""
    return _last_error

def calc_cost_jpy(input_tokens: int, output_tokens: int, model: str,
                   cache_read_tokens: int = 0, cache_creation_tokens: int = 0) -> float:
    """
    概算コストを円で返す（1USD=150円換算）。
    キャッシュ読み込みは通常入力の約0.1倍、キャッシュ書き込み（5分TTL）は約1.25倍で計算する。
    claude-sonnet-4-6: input $3/1M, output $15/1M
    claude-haiku-4-5:  input $1/1M, output $5/1M
    """
    rate = 150  # USD→JPY
    if 'haiku' in model:
        in_price, out_price = 1.00, 5.00
    else:
        in_price, out_price = 3.00, 15.00
    cost_usd = (
        input_tokens          / 1_000_000 * in_price
        + cache_read_tokens     / 1_000_000 * in_price * 0.1
        + cache_creation_tokens / 1_000_000 * in_price * 1.25
        + output_tokens         / 1_000_000 * out_price
    )
    return round(cost_usd * rate, 1)


def call_llm(client, prompt: str, schema: dict, provider: str, model: str,
             system: str | None = None, cache_system: bool = False) -> dict | list | None:
    """
    構造化出力でLLMを呼び出す。
    tool_use / function_calling → 失敗時はJSONフォールバックの順で試みる。
    トークン使用量・概算コストをグローバルカウンターに累積する。
    system/cache_systemは、複数回のバッチ呼び出しで変わらない部分（コードブックなど）を
    system側に分離しプロンプトキャッシュ（cache_control）を効かせたい場合に使う。
    """
    global _last_error
    last_reason = None
    # 1回目は再現性のためtemperature=0固定。空結果で失敗した場合、
    # 同一プロンプト×temperature=0では毎回同じ結果になり得るため、
    # リトライ時はtemperatureを上げて出力にばらつきを持たせる。
    # 呼び出し回数が多い方式C等でのトークン消費を抑えるため試行は2回までとする。
    temperatures = [0, 0.5]

    for attempt in range(len(temperatures)):
        try:
            temperature = temperatures[attempt]
            if provider == 'Anthropic':
                result, usage = _call_anthropic(client, prompt, schema, model, temperature, system, cache_system)
            elif provider == 'OpenAI':
                result, usage = _call_openai(client, prompt, schema, model, temperature, system)
            else:
                raise ValueError(f'未対応のプロバイダ: {provider}')

            if usage:
                _token_usage['input']          += usage.get('input', 0)
                _token_usage['output']         += usage.get('output', 0)
                _token_usage['cache_read']      += usage.get('cache_read', 0)
                _token_usage['cache_creation']  += usage.get('cache_creation', 0)
                _token_usage['cost_jpy']        += calc_cost_jpy(
                    usage.get('input', 0), usage.get('output', 0), model,
                    usage.get('cache_read', 0), usage.get('cache_creation', 0),
                )

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


def _call_anthropic(client, prompt: str, schema: dict, model: str, temperature: float = 0,
                     system: str | None = None, cache_system: bool = False) -> tuple:
    """
    Anthropic tool_use を使った構造化呼び出し。
    system指定時はsystemパラメータとして渡す。cache_system=Trueの場合、
    system全体にcache_control（プロンプトキャッシュ）を付与する
    （同じsystem文字列を使い回すバッチ処理で、2回目以降のコストを大幅に削減できる）。
    """
    tool = {
        'name': 'output_result',
        'description': 'プロンプトの指示に従って結果を構造化して返す',
        'input_schema': schema,
    }
    kwargs = dict(
        model=model,
        max_tokens=8192,
        temperature=temperature,
        tools=[tool],
        tool_choice={'type': 'tool', 'name': 'output_result'},
        messages=[{'role': 'user', 'content': prompt}],
    )
    if system:
        if cache_system:
            kwargs['system'] = [{'type': 'text', 'text': system, 'cache_control': {'type': 'ephemeral'}}]
        else:
            kwargs['system'] = system
    response = client.messages.create(**kwargs)
    usage = {
        'input':          response.usage.input_tokens,
        'output':         response.usage.output_tokens,
        'cache_read':     getattr(response.usage, 'cache_read_input_tokens', 0) or 0,
        'cache_creation': getattr(response.usage, 'cache_creation_input_tokens', 0) or 0,
    }
    if response.stop_reason == 'max_tokens':
        raise RuntimeError(
            f'出力がmax_tokens上限で打ち切られました（入力{usage["input"]}トークン/出力{usage["output"]}トークン）。'
            'データ量やコード数上限を減らすか、分割して再実行してください。'
        )

    tool_input = None
    text_parts = []
    for block in response.content:
        if block.type == 'tool_use' and block.name == 'output_result':
            tool_input = block.input
        elif block.type == 'text':
            text_parts.append(block.text)

    if tool_input:
        return tool_input, usage

    if text_parts:
        result = _parse_json_text('\n'.join(text_parts))
        if result:
            return result, usage

    # ここまで来た場合は空/無効な応答。原因診断のため詳細情報を添えて例外化する
    # （tool_use自体が来ていない場合はNoneのまま、空の引数で来た場合は{}が入る）
    preview = text_parts[0][:200] if text_parts else ''
    raise RuntimeError(
        f'空/無効な応答（stop_reason={response.stop_reason}, 出力{usage["output"]}トークン, '
        f'tool_use入力={tool_input!r}, テキスト先頭200字={preview!r}）'
    )


def _call_openai(client, prompt: str, schema: dict, model: str, temperature: float = 0,
                  system: str | None = None) -> tuple:
    """
    OpenAI function_calling を使った構造化呼び出し。
    OpenAI側はプロンプトキャッシュが自動適用されるため、cache_system相当の明示指定は不要。
    system指定時はsystemロールのメッセージとして渡すのみ。
    """
    tool = {
        'type': 'function',
        'function': {
            'name': 'output_result',
            'description': 'プロンプトの指示に従って結果を構造化して返す',
            'parameters': schema,
        }
    }
    messages = []
    if system:
        messages.append({'role': 'system', 'content': system})
    messages.append({'role': 'user', 'content': prompt})
    response = client.chat.completions.create(
        model=model,
        max_tokens=8192,
        temperature=temperature,
        tools=[tool],
        tool_choice={'type': 'function', 'function': {'name': 'output_result'}},
        messages=messages
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
