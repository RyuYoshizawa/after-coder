"""
llm_client.py
プロバイダ非依存のLLM呼び出しモジュール。
tool_use / function_calling / フォールバック（JSON文字列）を自動切替。
"""

import json
import re
import time


def call_llm(client, prompt: str, schema: dict, provider: str, model: str) -> dict | list | None:
    """
    構造化出力でLLMを呼び出す。
    tool_use / function_calling → 失敗時はJSONフォールバックの順で試みる。

    Args:
        client:   anthropic.Anthropic() または openai.OpenAI() のインスタンス
        prompt:   ユーザープロンプト
        schema:   出力スキーマ（JSONスキーマ形式）
        provider: 'Anthropic' または 'OpenAI'
        model:    使用するモデル名

    Returns:
        パース済みのdict/list、失敗時はNone
    """
    for attempt in range(3):
        try:
            if provider == 'Anthropic':
                result = _call_anthropic(client, prompt, schema, model)
            elif provider == 'OpenAI':
                result = _call_openai(client, prompt, schema, model)
            else:
                raise ValueError(f'未対応のプロバイダ: {provider}')

            if result is not None:
                return result

        except Exception as e:
            print(f'\n  [リトライ{attempt+1}] {type(e).__name__}: {e}')
            time.sleep(1)

    # 全て失敗した場合
    print('\n  [警告] 構造化出力に失敗しました。このバッチをスキップします。')
    return None


def _call_anthropic(client, prompt: str, schema: dict, model: str) -> dict | list | None:
    """Anthropic tool_use を使った構造化呼び出し"""
    tool = {
        'name': 'output_result',
        'description': 'プロンプトの指示に従って結果を構造化して返す',
        'input_schema': schema,
    }
    response = client.messages.create(
        model=model,
        max_tokens=4096,
        tools=[tool],
        tool_choice={'type': 'any'},
        messages=[{'role': 'user', 'content': prompt}]
    )
    # tool_useブロックを探す
    for block in response.content:
        if block.type == 'tool_use' and block.name == 'output_result':
            return block.input
    # tool_useがない場合はtextブロックからJSONをパース（フォールバック）
    for block in response.content:
        if block.type == 'text':
            result = _parse_json_text(block.text)
            if result is not None:
                return result
    return None


def _call_openai(client, prompt: str, schema: dict, model: str) -> dict | list | None:
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
        max_tokens=4096,
        tools=[tool],
        tool_choice='auto',
        messages=[{'role': 'user', 'content': prompt}]
    )
    msg = response.choices[0].message
    if msg.tool_calls:
        for tc in msg.tool_calls:
            if tc.function.name == 'output_result':
                return json.loads(tc.function.arguments)
    # フォールバック
    if msg.content:
        return _parse_json_text(msg.content)
    return None


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
