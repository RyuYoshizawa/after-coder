"""
アフターコーディング支援ツール - デモ版
Streamlit Webアプリ
"""

import streamlit as st
import anthropic
import hashlib
import json
import re
import time
import random
import io
import openpyxl
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from datetime import datetime
from llm_client import call_llm, make_client, reset_token_usage, get_token_usage, get_last_error

APP_DIR = Path(__file__).parent

# コードブック策定・編集は判断の質が重要なためSonnet固定。
# コーディング（既存コードブックへの分類作業）はサイドバーで選択可能（精度・価格比較のため）。
CODEBOOK_MODEL = 'claude-sonnet-4-6'
CODING_MODEL   = 'claude-haiku-4-5'  # コーディングモデルのデフォルト（サイドバー未選択時など）
CODING_MODEL_OPTIONS = {
    'Haiku 4.5（安い・高速）':   'claude-haiku-4-5',
    'Sonnet 4.6（高精度・高コスト）': 'claude-sonnet-4-6',
}

# リスクチェック項目。チェックした項目だけをコーディング時にAIへ問い合わせる（未チェックはコスト0）。
# key: 内部フィールド名（LLMスキーマ・result内で使用） / label: UI表示名
# char: Excel「回答別コーディング結果」の一文字見出し（チャンクB用） / hint: チェックの説明
RISK_CHECK_OPTIONS = [
    {'key': 'claim',    'label': 'クレーム',   'char': 'ク', 'hint': '強いクレーム'},
    {'key': 'personal', 'label': '個人情報',   'char': '個', 'hint': '個人名・メールアドレス・電話番号'},
    {'key': 'address',  'label': '住所情報',   'char': '住', 'hint': '詳細住所'},
    {'key': 'org',       'label': '団体情報',   'char': '団', 'hint': '学校名・病院名・施設名'},
    {'key': 'danger',    'label': '危険情報',   'char': '危', 'hint': '犯罪予告、自死予告、強い恨み'},
]

st.set_page_config(
    page_title='アフターコーディング支援ツール',
    page_icon='📊',
    layout='wide'
)
st.logo(str(APP_DIR / 'mj.png'), size='large')

# ── スタイル ──────────────────────────────────────────────────────
st.markdown("""
<style>
.main-title { font-size: 28px; font-weight: bold; color: #2E5C8A; margin-bottom: 0; }
.sub-title  { font-size: 14px; color: #666; margin-bottom: 24px; }
.step-label { font-size: 13px; font-weight: bold; color: #2E5C8A; }
.result-box { background: #f8f9fa; border-radius: 8px; padding: 16px; margin: 8px 0; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# 認証
# ══════════════════════════════════════════════════

USERS = {
    'starangler': 'QWEp12a23#',
    'mjguest': 'Amazonet1997',
    'KonomiSenda': 'Amazonet3944',
    'KunikoOkazaki': 'Mj3944',
}
ADMIN_USERNAME = 'admin'
ADMIN_PASSWORD = 'admin_pass_2026'  # 本番運用前に変更してください


def authenticate(username, password):
    if username == ADMIN_USERNAME:
        return password == ADMIN_PASSWORD
    return USERS.get(username) == password


st.session_state.setdefault('authenticated', False)
st.session_state.setdefault('username', None)
st.session_state.setdefault('history', [])
st.session_state.setdefault('history_counter', 0)
st.session_state.setdefault('active_history_id', None)
st.session_state.setdefault('texts_count', 0)


def render_login():
    st.markdown('<p class="main-title">👻 アフターコーディング支援ツール</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">ログインしてください</p>', unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        with st.form('login_form'):
            username = st.text_input('ユーザーID')
            password = st.text_input('パスワード', type='password')
            submitted = st.form_submit_button('ログイン', type='primary', width='stretch')
            if submitted:
                if authenticate(username, password):
                    st.session_state.authenticated = True
                    st.session_state.username = username
                    st.rerun()
                else:
                    st.error('ユーザーIDまたはパスワードが正しくありません')


if not st.session_state.authenticated:
    render_login()
    st.stop()


# ══════════════════════════════════════════════════
# LLM処理関数
# ══════════════════════════════════════════════════

def llm_generate_codebook(client, items, max_codes, q_name, data_context=''):
    text = '\n'.join(f'{x["id"]}: {x["text"]}' for x in items)
    context_str = f'\n【調査の背景・特徴】\n{data_context}\n' if data_context.strip() else ''
    prompt = f"""「{q_name}」の自由回答（{len(items)}件）からコードブックを作成してください。
{context_str}
【ルール】
- 中間カテゴリ7個前後（最大10個）
- カテゴリ内コード最大10個、総数{max_codes}個以内
- 1コード＝1主題
- コードID形式: CAT01.../C0101...（コードIDの接頭辞は英字1文字の「C」のみ。「CO」のように2文字以上にしない）

【回答リスト】
{text}"""

    schema = {
        'type': 'object',
        'properties': {
            'categories': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'cat_id':   {'type': 'string'},
                        'cat_name': {'type': 'string'},
                        'codes': {
                            'type': 'array',
                            'items': {
                                'type': 'object',
                                'properties': {
                                    'code_id':    {'type': 'string'},
                                    'code_name':  {'type': 'string'},
                                    'definition': {'type': 'string'},
                                },
                                'required': ['code_id', 'code_name', 'definition'],
                            }
                        }
                    },
                    'required': ['cat_id', 'cat_name', 'codes'],
                }
            }
        },
        'required': ['categories'],
    }
    return call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)


def llm_generate_codebook_topdown(client, data_context, q_name, max_codes):
    """（方式B用）実データを見る前に、調査の背景・特徴だけからカテゴリ骨格を生成"""
    prompt = f"""「{q_name}」の自由回答について、実際の回答データを見る前に、
調査の背景・特徴だけから想定されるコードブックの骨格（カテゴリと想定コード名）を作成してください。

【調査の背景・特徴】
{data_context}

【ルール】
- 中間カテゴリ7個前後（最大10個）
- カテゴリごとに想定コード名を列挙（1カテゴリあたり最大10個、総数{max_codes}個以内）
- コードIDは不要、コード名のみ列挙
- カテゴリIDはCAT01形式"""

    schema = {
        'type': 'object',
        'properties': {
            'categories': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'cat_id':         {'type': 'string'},
                        'cat_name':       {'type': 'string'},
                        'expected_codes': {'type': 'array', 'items': {'type': 'string'}},
                    },
                    'required': ['cat_id', 'cat_name', 'expected_codes'],
                }
            }
        },
        'required': ['categories'],
    }
    return call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)


def llm_elaborate_skeleton(client, skeleton, sample_items, max_codes, q_name, data_context=''):
    """（方式B用）トップダウン骨格を実データサンプルで具体化し、定義を付与したコードブックを確定"""
    skeleton_text = '\n'.join(
        f'{cat.get("cat_id", "")}: {cat.get("cat_name", "")} → ' + '、'.join(cat.get('expected_codes', []) or [])
        for cat in skeleton.get('categories', [])
    )
    sample_text = '\n'.join(f'{x["id"]}: {x["text"]}' for x in sample_items)
    context_str = f'\n【調査の背景・特徴】\n{data_context}\n' if data_context.strip() else ''
    prompt = f"""「{q_name}」のコードブックを確定してください。
以下は調査の背景から事前に想定したカテゴリ骨格です。骨格のカテゴリ構成（cat_id・cat_name）は維持しながら、
実際の回答サンプルを踏まえて各コードに定義を付与し、必要に応じてコードを追加・調整してください。

【カテゴリ骨格（トップダウンで抽出した想定コード）】
{skeleton_text}
{context_str}
【実際の回答サンプル（{len(sample_items)}件）】
{sample_text}

【ルール】
- 骨格のカテゴリ（cat_id・cat_name）は維持する
- 想定コードに定義を付与する（サンプルに現れない想定コードも残してよい）
- サンプルに現れる主題で骨格にないものは追加する（総数{max_codes}個以内）
- 1コード＝1主題
- コードID形式: CAT01.../C0101...（コードIDの接頭辞は英字1文字の「C」のみ。「CO」のように2文字以上にしない）"""

    schema = {
        'type': 'object',
        'properties': {
            'categories': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'cat_id':   {'type': 'string'},
                        'cat_name': {'type': 'string'},
                        'codes': {
                            'type': 'array',
                            'items': {
                                'type': 'object',
                                'properties': {
                                    'code_id':    {'type': 'string'},
                                    'code_name':  {'type': 'string'},
                                    'definition': {'type': 'string'},
                                },
                                'required': ['code_id', 'code_name', 'definition'],
                            }
                        }
                    },
                    'required': ['cat_id', 'cat_name', 'codes'],
                }
            }
        },
        'required': ['categories'],
    }
    return call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)


def llm_extract_topics(client, items, q_name, data_context=''):
    """（方式C用）回答バッチから主題（トピック）を網羅的に抽出"""
    text = '\n'.join(f'{x["id"]}: {x["text"]}' for x in items)
    context_str = f'\n【調査の背景・特徴】\n{data_context}\n' if data_context.strip() else ''
    prompt = f"""「{q_name}」の回答（{len(items)}件）に含まれる主題（トピック）をすべて抽出してください。
{context_str}
【ルール】
- 回答ごとに1つ以上の主題を短いフレーズで抽出する
- 似た内容は同じ表現にまとめてよい
- 網羅性を優先し、細かい主題も漏らさず列挙する

【回答リスト】
{text}"""

    schema = {
        'type': 'object',
        'properties': {
            'topics': {'type': 'array', 'items': {'type': 'string'}}
        },
        'required': ['topics'],
    }
    result = call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)
    if result and isinstance(result, dict):
        return result.get('topics', [])
    return []


def llm_reduce_topics(client, topics, q_name, data_context=''):
    """（方式C用）主題リストのチャンクを、同義・類似表現をまとめて簡潔なユニークリストに整理する"""
    topics_text = '\n'.join(f'- {t}' for t in topics)
    context_str = f'\n【調査の背景・特徴】\n{data_context}\n' if data_context.strip() else ''
    prompt = f"""「{q_name}」の回答から抽出された主題リスト（{len(topics)}件）を整理してください。
{context_str}
【主題リスト】
{topics_text}

【ルール】
- 同義・類似する主題はひとつにまとめる
- 表現はできるだけ簡潔な短いフレーズにする
- 異なる主題の情報を失わないよう、まとめすぎない"""

    schema = {
        'type': 'object',
        'properties': {
            'topics': {'type': 'array', 'items': {'type': 'string'}}
        },
        'required': ['topics'],
    }
    result = call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)
    if result and isinstance(result, dict):
        return result.get('topics', [])
    return topics  # 失敗時は元のチャンクをそのまま返し情報を失わない


def llm_consolidate_topics(client, topics, max_codes, q_name, data_context=''):
    """（方式C用）全件から蓄積した主題リストを統合し、定義・キーワード付きのコードブックを確定"""
    topics_text = '\n'.join(f'- {t}' for t in topics)
    context_str = f'\n【調査の背景・特徴】\n{data_context}\n' if data_context.strip() else ''
    prompt = f"""「{q_name}」の全回答から抽出された主題リスト（{len(topics)}件、類似表現含む）を統合し、
最終的なコードブックを確定してください。
{context_str}
【主題リスト】
{topics_text}

【ルール】
- 類似・重複する主題を統合し、中間カテゴリ7個前後（最大10個）に整理する
- カテゴリ内コード最大10個、総数{max_codes}個以内
- 1コード＝1主題
- 各コードに定義と、判定に役立つキーワードを2〜5個付与する
- コードID形式: CAT01.../C0101...（コードIDの接頭辞は英字1文字の「C」のみ。「CO」のように2文字以上にしない）"""

    schema = {
        'type': 'object',
        'properties': {
            'categories': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'cat_id':   {'type': 'string'},
                        'cat_name': {'type': 'string'},
                        'codes': {
                            'type': 'array',
                            'items': {
                                'type': 'object',
                                'properties': {
                                    'code_id':    {'type': 'string'},
                                    'code_name':  {'type': 'string'},
                                    'definition': {'type': 'string'},
                                    'keywords':   {'type': 'array', 'items': {'type': 'string'}},
                                },
                                'required': ['code_id', 'code_name', 'definition'],
                            }
                        }
                    },
                    'required': ['cat_id', 'cat_name', 'codes'],
                }
            }
        },
        'required': ['categories'],
    }
    return call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)


def llm_detect_new_codes(client, items, codebook, max_codes, q_name):
    text = '\n'.join(f'{x["id"]}: {x["text"]}' for x in items)
    existing = '\n'.join(
        f'{c["code_id"]}({cat["cat_id"]}:{cat["cat_name"]}):{c["code_name"]}'
        for cat in codebook.get('categories', [])
        for c in cat.get('codes', [])
    )
    cat_list = '\n'.join(
        f'{cat["cat_id"]}: {cat["cat_name"]}'
        for cat in codebook.get('categories', [])
    )
    cur = sum(len(c['codes']) for c in codebook.get('categories', []))
    prompt = f"""「{q_name}」の回答（{len(items)}件）に既存コードで対応できない新主題がありますか？

【既存カテゴリ一覧】
{cat_list}

【既存コード（{cur}個、残り{max_codes-cur}個）】
{existing}

【回答】
{text}

新主題がある場合のみ返してください。なければnew_codesを空配列に。
cat_idは必ず既存カテゴリ一覧のCAT01形式を使うこと。
code_idは既存コードと同じC0101形式で採番すること（接頭辞は英字1文字の「C」のみ。「CO」のように2文字以上にしない）。"""

    schema = {
        'type': 'object',
        'properties': {
            'new_codes': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'cat_id':    {'type': 'string'},
                        'code_id':   {'type': 'string'},
                        'code_name': {'type': 'string'},
                        'definition':{'type': 'string'},
                    },
                    'required': ['cat_id', 'code_id', 'code_name', 'definition'],
                }
            }
        },
        'required': ['new_codes'],
    }
    result = call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)
    if result is None:
        return []
    return result.get('new_codes', [])


def llm_code_batch(client, items, codes, q_name, model=CODING_MODEL, enabled_risks=None):
    """
    コーディング本体。同じコードブック（system側）を1回のコーディング実行中に何十回も
    使い回すため、コードブック・ルールをsystemに分離しcache_control（プロンプトキャッシュ）を
    効かせ、2回目以降のバッチでコードブック分のコストを大幅に削減する。
    modelはサイドバーの「コーディングモデル」選択に従う（デフォルトはHaiku=CODING_MODEL）。
    精度・価格を比較したい場合は、同じコードブックのまま方式を変えて2回実行し、
    作業履歴で比較する使い方を想定している。
    enabled_risksはサイドバーの「リスクチェック」でチェックされた項目のkeyのリスト。
    未チェックの項目はプロンプト・スキーマに一切含めない（AIへの問い合わせ自体をしない＝コスト増なし）。
    """
    enabled_risks = enabled_risks or []
    risk_opts = [o for o in RISK_CHECK_OPTIONS if o['key'] in enabled_risks]

    code_list = '\n'.join(
        f'{c["code_id"]}（{c["cat_name"]}）: {c["code_name"]} / {c["definition"][:25]}'
        for c in codes
    )
    risk_rule = ''
    if risk_opts:
        risk_lines = '\n'.join(f"- {o['label']}: {o['hint']}" for o in risk_opts)
        risk_rule = f"""

【リスクチェック】次の該当有無も回答ごとに判定してください（該当すればtrue、しなければfalse）
{risk_lines}"""

    system_prompt = f"""「{q_name}」の回答にコードブックに基づいてコーディングしてください。

【コードブック】
{code_list}

【ルール】
- 1回答に複数コード付与可
- 該当なしはcodesを空配列
- sentimentは positive/negative/neutral{risk_rule}"""

    items_text = '\n'.join(f'{x["id"]}: {x["text"]}' for x in items)
    prompt = f"""【回答】
{items_text}"""

    result_properties = {
        'id':        {'type': 'string'},
        'codes':     {'type': 'array', 'items': {'type': 'string'}},
        'sentiment': {'type': 'string', 'enum': ['positive','negative','neutral']},
    }
    for o in risk_opts:
        result_properties[o['key']] = {'type': 'boolean'}

    schema = {
        'type': 'object',
        'properties': {
            'results': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': result_properties,
                    'required': ['id', 'codes', 'sentiment'],
                }
            }
        },
        'required': ['results'],
    }
    result = call_llm(client, prompt, schema, 'Anthropic', model,
                       system=system_prompt, cache_system=True)
    if result and isinstance(result, dict):
        return result.get('results', [])
    return []


def llm_edit_codebook(client, codebook, instruction, q_name):
    """
    コードブック編集機能：チャット指示に基づき、統合・改名・再定義・追加・削除などの編集を行う。
    生データは参照しない（新規コードの発見は行わず、既存コードブックの整理のみ）。
    """
    codebook_text = json.dumps(codebook, ensure_ascii=False)
    prompt = f"""「{q_name}」のコードブックを、次の指示に従って編集してください。

【現在のコードブック】
{codebook_text}

【編集指示】
{instruction}

【ルール】
- 指示された編集（統合・改名・再定義・コードの追加・削除など）のみを行う
- 指示にない部分はそのまま維持する
- 生データは参照できないため、指示にない新規コードの発見は行わない
- コードID・カテゴリIDの形式は維持する（新規追加時はCAT.../C0101形式で採番する。コードIDの接頭辞は英字1文字の「C」のみ。「CO」のように2文字以上にしない）
- 編集後のコードブック全体を返す"""

    schema = {
        'type': 'object',
        'properties': {
            'categories': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'cat_id':   {'type': 'string'},
                        'cat_name': {'type': 'string'},
                        'codes': {
                            'type': 'array',
                            'items': {
                                'type': 'object',
                                'properties': {
                                    'code_id':    {'type': 'string'},
                                    'code_name':  {'type': 'string'},
                                    'definition': {'type': 'string'},
                                    'keywords':   {'type': 'array', 'items': {'type': 'string'}},
                                },
                                'required': ['code_id', 'code_name', 'definition'],
                            }
                        }
                    },
                    'required': ['cat_id', 'cat_name', 'codes'],
                }
            }
        },
        'required': ['categories'],
    }
    return call_llm(client, prompt, schema, 'Anthropic', CODEBOOK_MODEL)
# ══════════════════════════════════════════════════
# 集計・Excel出力関数
# ══════════════════════════════════════════════════

def aggregate_results(codes, results, total, risk_keys=None):
    """
    コード別GT集計とセンチメント集計・非該当（どのコードも付与されなかった回答）件数、
    リスクチェック（risk_keysで有効な項目のみ）の該当件数を算出
    """
    risk_keys   = risk_keys or []
    code_counts = {c['code_id']: 0 for c in codes}
    sent_counts = {'positive': 0, 'negative': 0, 'neutral': 0}
    risk_counts = {k: 0 for k in risk_keys}
    result_map  = {r['id']: r for r in results}
    unassigned  = 0

    for rid, res in result_map.items():
        sent = res.get('sentiment', 'neutral')
        if sent in sent_counts:
            sent_counts[sent] += 1
        res_codes = res.get('codes', [])
        if not res_codes:
            unassigned += 1
        for cid in res_codes:
            if cid in code_counts:
                code_counts[cid] += 1
        for k in risk_keys:
            if res.get(k):
                risk_counts[k] += 1

    gt = []
    for cat in codes:
        cnt = code_counts.get(cat['code_id'], 0)
        pct = cnt / total * 100 if total > 0 else 0
        gt.append({
            'cat_id':    cat['cat_id'],
            'cat_name':  cat['cat_name'],
            'code_id':   cat['code_id'],
            'code_name': cat['code_name'],
            'count':     cnt,
            'pct':       round(pct, 1),
            'definition':cat.get('definition', ''),
        })
    gt.sort(key=lambda x: x['count'], reverse=True)
    return gt, sent_counts, unassigned, risk_counts


CODEBOOK_CSV_COLUMNS = ['カテゴリID', 'カテゴリ名', 'コードID', 'コード名', '定義', 'キーワード']


def _codebook_rows(codebook, gt_by_code=None, include_stats=False):
    """コードブックをカテゴリ→コードの行データ（dictのリスト）に変換する共通処理"""
    gt_by_code = gt_by_code or {}
    rows = []
    for cat in codebook.get('categories', []):
        for c in cat.get('codes', []):
            keywords = c.get('keywords', [])
            if isinstance(keywords, list):
                keywords = '; '.join(keywords)
            row = {
                'カテゴリID': cat.get('cat_id', ''),
                'カテゴリ名': cat.get('cat_name', ''),
                'コードID':   c.get('code_id', ''),
                'コード名':   c.get('code_name', ''),
                '定義':       c.get('definition', ''),
                'キーワード': keywords or '',
            }
            if include_stats:
                stat = gt_by_code.get(c.get('code_id'), {})
                row = {'件数': stat.get('count', 0), '出現率(%)': stat.get('pct', 0.0), **row}
            rows.append(row)
    return rows


def render_codebook_structure(codebook, key=None):
    """
    コードブックの構造（カテゴリID・カテゴリ名・コードID・コード名・定義・キーワード）のみを、
    折りたたまず常時表示する。編集直後もその場で最新の内容が確認できる。
    セルをダブルクリックするとテキストを全文選択・コピーできる（st.data_editorを表示専用に使用。
    ここでの編集内容は保存されない＝実際のコードブックには反映されない）。
    表右上のツールバーからCSVダウンロードでき、そのCSVは「既存のコードブックを使用」で再読み込みできる。
    """
    import pandas as pd
    rows = _codebook_rows(codebook)
    n_cats = len(codebook.get('categories', []))
    st.caption(f'カテゴリ{n_cats}／コード{len(rows)}　※セルをダブルクリックするとテキストをコピーできます（ここでの編集内容は保存されません）')
    st.data_editor(pd.DataFrame(rows), width='stretch', hide_index=True, key=key)


def render_code_list_table(codebook, gt_by_code=None, expanded=False, key=None):
    """
    コーディング結果に基づく「コード一覧集計」を折りたたみ表示する。
    列順は 件数・出現率(%)・カテゴリID・カテゴリ名・コードID・コード名・定義・キーワード。
    gt_by_codeを渡さない、または未コーディングのコードは件数・出現率とも0として表示する。
    セルをダブルクリックするとテキストを全文選択・コピーできる（render_codebook_structureと同様）。
    """
    import pandas as pd
    rows = _codebook_rows(codebook, gt_by_code, include_stats=True)
    n_cats = len(codebook.get('categories', []))
    with st.expander(f'📋 コード一覧集計（カテゴリ{n_cats}／コード{len(rows)}）', expanded=expanded):
        st.caption('※セルをダブルクリックするとテキストをコピーできます（ここでの編集内容は保存されません）')
        st.data_editor(pd.DataFrame(rows), width='stretch', hide_index=True, key=key)


def parse_codebook_csv(file_bytes):
    """
    render_codebook_structure/render_code_list_tableの表からダウンロードしたCSV
    （カテゴリID,カテゴリ名,コードID,コード名,定義,キーワード。件数・出現率列があっても無視する）を
    コードブック構造（{'categories': [...]}）に復元する
    """
    import pandas as pd
    for enc in ('utf-8-sig', 'cp932'):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, dtype=str, keep_default_na=False)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError('文字コードを判定できませんでした（UTF-8またはShift-JISで保存し直してください）')

    missing = [c for c in CODEBOOK_CSV_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f'必要な列が見つかりません: {", ".join(missing)}')

    categories = {}
    for _, row in df.iterrows():
        cat_id = row['カテゴリID'].strip()
        if not cat_id:
            continue
        if cat_id not in categories:
            categories[cat_id] = {'cat_id': cat_id, 'cat_name': row['カテゴリ名'].strip(), 'codes': []}
        keywords = [k.strip() for k in row['キーワード'].split(';') if k.strip()]
        categories[cat_id]['codes'].append({
            'code_id':    row['コードID'].strip(),
            'code_name':  row['コード名'].strip(),
            'definition': row['定義'].strip(),
            'keywords':   keywords,
        })
    return {'categories': list(categories.values())}


def create_excel(q_name, gt, sent_counts, total, results, items, codes, unassigned=0,
                  risk_counts=None, enabled_risks=None):
    """
    ローカル版仮集計シートと同じレイアウトでExcelを生成。
    リスクチェック（risk_counts/enabled_risks）は「特記情報集計」への項目追加と、
    「回答別コーディング結果」の一文字フラグ列（非該当＋有効化されたリスク項目）に反映する。
    """
    risk_counts   = risk_counts or {}
    enabled_risks = enabled_risks or []
    risk_opts     = [o for o in RISK_CHECK_OPTIONS if o['key'] in enabled_risks]

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = '集計レポート'
    ws.sheet_view.showGridLines = False

    THIN      = Side(style='thin', color='C0C0C0')
    BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    HDR_FILL  = PatternFill('solid', start_color='2E5C8A', end_color='2E5C8A')
    HDR_FONT  = Font(name='Meiryo UI', bold=True, color='FFFFFF', size=10)
    SUB_FONT  = Font(name='Meiryo UI', bold=True, size=10)
    DATA_FONT = Font(name='Meiryo UI', size=10)
    HIGH_FILL = PatternFill('solid', start_color='E2EFDA', end_color='E2EFDA')
    LOW_FILL  = PatternFill('solid', start_color='FCE4D6', end_color='FCE4D6')
    FLAG_FILL = PatternFill('solid', start_color='BDD7EE', end_color='BDD7EE')

    # カテゴリ別カラー：画面の縦棒グラフ（Plotlyデフォルト配色、カテゴリ出現率の多い順に割当）と
    # 同じ色を使う。「中間カテゴリID」行・コード列見出し行はフル彩度、他の行はその淡色版にする。
    PLOTLY_COLORS = ['636EFA','EF553B','00CC96','AB63FA','FFA15A',
                      '19D3F3','FF6692','B6E880','FF97FF','FECB52']

    def _lighten_hex(hex_color, factor=0.65):
        """指定した割合(0〜1)だけ白に近づけた淡い色を返す"""
        r = int(hex_color[0:2], 16); g = int(hex_color[2:4], 16); b = int(hex_color[4:6], 16)
        r = round(r + (255 - r) * factor)
        g = round(g + (255 - g) * factor)
        b = round(b + (255 - b) * factor)
        return f'{r:02X}{g:02X}{b:02X}'

    cat_total = {}
    for g in gt:
        cat_total[g['cat_id']] = cat_total.get(g['cat_id'], 0) + g['count']
    for c in codes:
        cat_total.setdefault(c['cat_id'], 0)
    cat_order      = sorted(cat_total, key=lambda cid: -cat_total[cid])
    cat_color_full = {cid: PLOTLY_COLORS[i % len(PLOTLY_COLORS)] for i, cid in enumerate(cat_order)}
    cat_color_pale = {cid: _lighten_hex(col) for cid, col in cat_color_full.items()}

    def hdr(r, c, v):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font=HDR_FONT; cell.fill=HDR_FILL; cell.border=BORDER
        return cell

    def lbl(r, c, v):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font=HDR_FONT; cell.fill=HDR_FILL; cell.border=BORDER
        return cell

    def dat(r, c, v, fill=None):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font=DATA_FONT; cell.border=BORDER
        if fill: cell.fill = fill
        return cell

    # ── タイトル ──────────────────────────────────────────────────
    ws.cell(row=1, column=1, value=f'集計レポート：{q_name}').font = Font(
        name='Meiryo UI', bold=True, size=14, color='2E5C8A')
    ws.cell(row=2, column=1,
            value=f'集計日時: {datetime.now().strftime("%Y/%m/%d %H:%M")}  有効回答数: {total}件').font = Font(
        name='Meiryo UI', size=10, color='808080')

    # 回答別コーディング結果の固定列：回答ID・回答テキスト・センチメント＋フラグ列（非該当＋有効なリスクチェック項目）
    FLAG_COLUMNS = [{'key': None, 'char': '非'}] + [{'key': o['key'], 'char': o['char']} for o in risk_opts]
    FIXED_N    = 3 + len(FLAG_COLUMNS)
    CODE_START = FIXED_N + 1

    # ── 特記情報集計（センチメント＋非該当＋リスクチェック） ────────────
    ws.cell(row=4, column=1, value='■ 特記情報集計').font = SUB_FONT
    sent_order = [('positive','ポジティブ'), ('negative','ネガティブ'), ('neutral','ニュートラル')]
    for i, (sent, label) in enumerate(sent_order):
        cnt = sent_counts[sent]
        pct = cnt / total * 100 if total > 0 else 0
        hdr(5, 4+i*2, label)
        dat(5, 5+i*2, cnt)
        hdr(6, 4+i*2, '%')
        dat(6, 5+i*2, round(pct, 1))

    # 非該当（どのコードも付与されなかった回答）
    un_pct = unassigned / total * 100 if total > 0 else 0
    hdr(5, 4+len(sent_order)*2, '非該当（コードなし）')
    dat(5, 5+len(sent_order)*2, unassigned)
    hdr(6, 4+len(sent_order)*2, '%')
    dat(6, 5+len(sent_order)*2, round(un_pct, 1))

    # リスクチェック（有効化された項目のみ）
    risk_base_col = 4 + (len(sent_order)+1)*2
    for i, o in enumerate(risk_opts):
        cnt = risk_counts.get(o['key'], 0)
        pct = cnt / total * 100 if total > 0 else 0
        col = risk_base_col + i*2
        hdr(5, col, o['label'])
        dat(5, col+1, cnt)
        hdr(6, col, '%')
        dat(6, col+1, round(pct, 1))

    # ── GT集計（転置レイアウト） ──────────────────────────────────
    GT_START  = 8
    ws.cell(row=GT_START, column=1, value='■ コード別GT集計').font = SUB_FONT

    gt_labels = [
        'カテゴリ出現数',
        '中間カテゴリID',
        '中間カテゴリ名',
        'コードID',
        'コード名',
        '出現件数',
        '出現率(%)',
        '定義',
        'ポジティブ件数',
        'ネガティブ件数',
        'ニュートラル件数',
    ]
    CAT_ID_ROW_INDEX = 1  # gt_labelsのうち「中間カテゴリID」の位置＝フル彩度で塗る行
    for i, label in enumerate(gt_labels):
        r = GT_START + 1 + i
        lbl(r, 1, label)
        ws.row_dimensions[r].height = 18

    # カテゴリ出現数を集計
    result_map  = {r['id']: r for r in results}
    cat_counts  = {}
    code_sent   = {c['code_id']: {'positive':0,'negative':0,'neutral':0} for c in codes}
    for res in results:
        sent = res.get('sentiment', 'neutral')
        assigned_cats = set()
        for cid in res.get('codes', []):
            code = next((c for c in codes if c['code_id']==cid), None)
            if code:
                cat = code['cat_id']
                if cat not in assigned_cats:
                    cat_counts[cat] = cat_counts.get(cat, 0) + 1
                    assigned_cats.add(cat)
                if sent in code_sent[cid]:
                    code_sent[cid][sent] += 1

    # コードデータ書き込み
    for ci, code in enumerate(codes):
        col  = CODE_START + ci
        cnt  = next((r['count'] for r in gt if r['code_id']==code['code_id']), 0)
        pct  = cnt / total * 100 if total > 0 else 0
        cs   = code_sent.get(code['code_id'], {'positive':0,'negative':0,'neutral':0})
        full = cat_color_full.get(code['cat_id'], 'FFFFFF')
        pale = cat_color_pale.get(code['cat_id'], 'FFFFFF')
        vals = [
            cat_counts.get(code['cat_id'], 0),
            code['cat_id'],
            code['cat_name'],
            code['code_id'],
            code['code_name'],
            cnt,
            round(pct, 1),
            code.get('definition', ''),
            cs['positive'],
            cs['negative'],
            cs['neutral'],
        ]
        for ri, val in enumerate(vals):
            r = GT_START + 1 + ri
            c = ws.cell(row=r, column=col, value=val)
            c.font=DATA_FONT; c.border=BORDER
            c.alignment = Alignment(wrap_text=True, vertical='top')
            row_color = full if ri == CAT_ID_ROW_INDEX else pale
            c.fill = PatternFill('solid', start_color=row_color, end_color=row_color)
            if ri == 6:  # 出現率
                if pct >= 20:
                    c.fill = HIGH_FILL
                elif pct <= 2:
                    c.fill = LOW_FILL
                    c.font = Font(name='Meiryo UI', size=10, color='FF0000', bold=True)
        ws.column_dimensions[get_column_letter(col)].width = 12

    # ── 回答別コーディング結果 ────────────────────────────────────
    RESULT_START = GT_START + len(gt_labels) + 2
    ws.cell(row=RESULT_START, column=1, value='■ 回答別コーディング結果').font = SUB_FONT

    hdr_row = RESULT_START + 1
    for ci, h in enumerate(['回答ID', '回答テキスト', 'センチメント']):
        hdr(hdr_row, 1+ci, h)

    # フラグ列見出し（非該当＋有効化されたリスクチェック項目、一文字見出し）
    for ci, fc in enumerate(FLAG_COLUMNS):
        hdr(hdr_row, 4+ci, fc['char'])
        ws.column_dimensions[get_column_letter(4+ci)].width = 4

    for ci, code in enumerate(codes):
        col  = CODE_START + ci
        full = cat_color_full.get(code['cat_id'], 'FFFFFF')
        c = ws.cell(row=hdr_row, column=col, value=code['code_id'])
        c.font = Font(name='Meiryo UI', bold=True, size=10, color='000000')
        c.fill = PatternFill('solid', start_color=full, end_color=full); c.border=BORDER

    item_map = {item['id']: item['text'] for item in items}
    for ri, res in enumerate(results):
        r        = hdr_row + 1 + ri
        rid      = res.get('id', '')
        text     = item_map.get(rid, '')[:50]
        sent     = res.get('sentiment', '')
        assigned = res.get('codes', [])
        for ci, val in enumerate([rid, text, sent]):
            dat(r, 1+ci, val)
        for ci, fc in enumerate(FLAG_COLUMNS):
            flagged = (not assigned) if fc['key'] is None else bool(res.get(fc['key']))
            c = ws.cell(row=r, column=4+ci, value=1 if flagged else '')
            c.font=DATA_FONT; c.border=BORDER
            if flagged: c.fill = FLAG_FILL
        for ci, code in enumerate(codes):
            col  = CODE_START + ci
            flag = 1 if code['code_id'] in assigned else 0
            pale = cat_color_pale.get(code['cat_id'], 'FFFFFF')
            c = ws.cell(row=r, column=col, value=flag if flag else '')
            c.font=DATA_FONT; c.border=BORDER
            c.fill = PatternFill('solid', start_color=pale, end_color=pale)
            if flag: c.fill = FLAG_FILL

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 12
    ws.freeze_panes = f'A{hdr_row+1}'

    # オートフィルター（見出し行〜最終回答行、全列に設定）
    last_col = CODE_START + len(codes) - 1 if codes else FIXED_N
    ws.auto_filter.ref = f'A{hdr_row}:{get_column_letter(last_col)}{hdr_row + len(results)}'

    # ══════════════════════════════════════════════════
    # 集計ダイジェストシート
    # ══════════════════════════════════════════════════
    ws2 = wb.create_sheet('集計ダイジェスト')
    ws2.sheet_view.showGridLines = False

    # タイトル
    ws2.cell(row=1, column=1, value=f'After Coder 集計ダイジェスト：{q_name}').font = Font(
        name='Meiryo UI', bold=True, size=14, color='2E5C8A')
    ws2.cell(row=2, column=1,
             value=f'集計日時: {datetime.now().strftime("%Y/%m/%d %H:%M")}  有効回答数: {total}件').font = Font(
        name='Meiryo UI', size=10, color='808080')

    # 特記情報集計
    ws2.cell(row=4, column=1, value='■ 特記情報集計').font = SUB_FONT
    sent_order2 = [('positive','ポジティブ'),('negative','ネガティブ'),('neutral','ニュートラル')]
    for i, (sent, label) in enumerate(sent_order2):
        cnt = sent_counts[sent]
        pct = cnt / total * 100 if total > 0 else 0
        c = ws2.cell(row=5, column=1+i*2, value=label)
        c.font=HDR_FONT; c.fill=HDR_FILL; c.border=BORDER
        c2 = ws2.cell(row=5, column=2+i*2, value=cnt)
        c2.font=DATA_FONT; c2.border=BORDER
        c3 = ws2.cell(row=6, column=1+i*2, value='%')
        c3.font=HDR_FONT; c3.fill=HDR_FILL; c3.border=BORDER
        c4 = ws2.cell(row=6, column=2+i*2, value=round(pct,1))
        c4.font=DATA_FONT; c4.border=BORDER

    # 非該当（どのコードも付与されなかった回答）
    un_pct = unassigned / total * 100 if total > 0 else 0
    c = ws2.cell(row=5, column=1+len(sent_order2)*2, value='非該当（コードなし）')
    c.font=HDR_FONT; c.fill=HDR_FILL; c.border=BORDER
    c2 = ws2.cell(row=5, column=2+len(sent_order2)*2, value=unassigned)
    c2.font=DATA_FONT; c2.border=BORDER
    c3 = ws2.cell(row=6, column=1+len(sent_order2)*2, value='%')
    c3.font=HDR_FONT; c3.fill=HDR_FILL; c3.border=BORDER
    c4 = ws2.cell(row=6, column=2+len(sent_order2)*2, value=round(un_pct,1))
    c4.font=DATA_FONT; c4.border=BORDER

    # リスクチェック（有効化された項目のみ）
    risk_base_col2 = 1 + (len(sent_order2)+1)*2
    for i, o in enumerate(risk_opts):
        cnt = risk_counts.get(o['key'], 0)
        pct = cnt / total * 100 if total > 0 else 0
        col = risk_base_col2 + i*2
        c = ws2.cell(row=5, column=col, value=o['label'])
        c.font=HDR_FONT; c.fill=HDR_FILL; c.border=BORDER
        c2 = ws2.cell(row=5, column=col+1, value=cnt)
        c2.font=DATA_FONT; c2.border=BORDER
        c3 = ws2.cell(row=6, column=col, value='%')
        c3.font=HDR_FONT; c3.fill=HDR_FILL; c3.border=BORDER
        c4 = ws2.cell(row=6, column=col+1, value=round(pct,1))
        c4.font=DATA_FONT; c4.border=BORDER

    # GT集計（行方向・出現率順）
    ws2.cell(row=8, column=1, value='■ コード別GT集計').font = SUB_FONT
    gt_headers = ['カテゴリID', 'カテゴリ名', 'コードID', 'コード名', '出現件数', '出現率(%)', '定義', 'キーワード']
    for ci, h in enumerate(gt_headers):
        c = ws2.cell(row=9, column=ci+1, value=h)
        c.font=HDR_FONT; c.fill=HDR_FILL; c.border=BORDER

    code_kw_map = {}
    for cd in codes:
        kw = cd.get('keywords', [])
        if isinstance(kw, list):
            kw = '; '.join(kw)
        code_kw_map[cd['code_id']] = kw or ''

    # 出現率順にソート
    gt_sorted = sorted(gt, key=lambda x: x['count'], reverse=True)
    for ri, row_data in enumerate(gt_sorted):
        r    = 10 + ri
        pct  = row_data['pct']
        pale = cat_color_pale.get(row_data['cat_id'], 'FFFFFF')
        vals = [
            row_data['cat_id'],
            row_data['cat_name'],
            row_data['code_id'],
            row_data['code_name'],
            row_data['count'],
            pct,
            row_data['definition'],
            code_kw_map.get(row_data['code_id'], ''),
        ]
        for ci, val in enumerate(vals):
            c = ws2.cell(row=r, column=ci+1, value=val)
            c.font=DATA_FONT; c.border=BORDER
            c.fill = PatternFill('solid', start_color=pale, end_color=pale)
            if ci == 5:  # 出現率列
                if pct >= 20:
                    c.fill = HIGH_FILL
                elif pct <= 2:
                    c.fill = LOW_FILL
                    c.font = Font(name='Meiryo UI', size=10, color='FF0000', bold=True)

    # 列幅
    for ci, w in enumerate([10, 22, 10, 24, 10, 10, 40, 30]):
        ws2.column_dimensions[get_column_letter(ci+1)].width = w
    ws2.freeze_panes = 'A10'

    # 棒グラフ（画面の「コード別GT集計」と同じ配色。データポイントごとに色を指定）
    if gt_sorted:
        chart_row_start = 10
        chart_row_end   = 10 + len(gt_sorted) - 1
        chart = BarChart()
        chart.type  = 'col'
        chart.title = 'コード別出現率(%)'
        chart.y_axis.title = '出現率(%)'
        data = Reference(ws2, min_col=6, min_row=9, max_row=chart_row_end)
        cats = Reference(ws2, min_col=4, min_row=chart_row_start, max_row=chart_row_end)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        series = chart.series[0]
        series.data_points = [
            DataPoint(idx=i, spPr=GraphicalProperties(
                solidFill=cat_color_full.get(row_data['cat_id'], 'FFFFFF')))
            for i, row_data in enumerate(gt_sorted)
        ]
        chart.width  = 24
        chart.height = 10
        ws2.add_chart(chart, f'A{chart_row_end + 3}')

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ══════════════════════════════════════════════════
# コードブック策定方式（A/B/C）
# ══════════════════════════════════════════════════

CODEBOOK_MODES = [
    {'code': 'A',        'label': '方式A：標準'},
    {'code': 'B',        'label': '方式B：トップダウン＋ボトムアップ（推奨）'},
    {'code': 'C1',       'label': '方式C1：全件精査（1/4抽出）'},
    {'code': 'C2',       'label': '方式C2：全件精査（ハーフ抽出）'},
    {'code': 'C3',       'label': '方式C3：全件精査（フル抽出）'},
    {'code': 'EXISTING', 'label': '既存のコードブックを使用（アップロード）'},
]
CODEBOOK_MODE_LABELS  = [m['label'] for m in CODEBOOK_MODES]
CODEBOOK_MODE_CODE    = {m['label']: m['code'] for m in CODEBOOK_MODES}
CODEBOOK_MODE_C_RATIO = {'C1': 0.25, 'C2': 0.5, 'C3': 1.0}


def _diff_detect_loop(client, codebook, remaining, max_codes, q_name, progress_bar, p_start=0.05, p_end=0.30):
    """既存コードブックに対し残りサンプルをバッチ処理し新規コードを差分検出する（方式A・Bで共用）"""
    total_batches = max(len(remaining) // 20, 1)
    round_no = 2
    while remaining:
        batch     = remaining[:20]
        remaining = remaining[20:]
        new_list  = llm_detect_new_codes(client, batch, codebook, max_codes, q_name)
        if new_list:
            for nc in new_list:
                for cat in codebook['categories']:
                    if cat['cat_id'] == nc.get('cat_id'):
                        cat['codes'].append({
                            'code_id':    nc.get('code_id', ''),
                            'code_name':  nc.get('code_name', ''),
                            'definition': nc.get('definition', ''),
                        })
                        break
        cur = sum(len(c['codes']) for c in codebook.get('categories', []))
        if cur >= max_codes:
            break
        progress_bar.progress(min(p_start + (p_end - p_start) * (round_no / total_batches), p_end))
        round_no += 1
    return codebook


def _build_codebook_a(client, all_items, max_codes, q_name, data_context, progress_bar):
    """方式A：標準 - 少量サンプルでコードブックを生成し差分検出"""
    sample_1 = all_items[:30]
    codebook = llm_generate_codebook(client, sample_1, max_codes, q_name, data_context)
    if not codebook:
        return None
    return _diff_detect_loop(client, codebook, all_items[30:], max_codes, q_name, progress_bar)


def _build_codebook_b(client, all_items, max_codes, q_name, data_context, progress_bar, status_text):
    """方式B：トップダウン＋ボトムアップ - 骨格を先に生成し実データで具体化・差分検出"""
    if not data_context.strip():
        status_text.markdown('**Step 1/3** 「分析データの特徴」が未入力のため方式Aで生成します...')
        return _build_codebook_a(client, all_items, max_codes, q_name, data_context, progress_bar)

    status_text.markdown('**Step 1/3** トップダウンで骨格を生成中...')
    skeleton = llm_generate_codebook_topdown(client, data_context, q_name, max_codes)
    if not skeleton:
        return None
    progress_bar.progress(0.10)

    status_text.markdown('**Step 1/3** 実データで骨格を具体化中...')
    codebook = llm_elaborate_skeleton(client, skeleton, all_items[:30], max_codes, q_name, data_context)
    if not codebook:
        return None
    progress_bar.progress(0.20)

    return _diff_detect_loop(client, codebook, all_items[30:], max_codes, q_name, progress_bar, p_start=0.20, p_end=0.30)


def _topics_cache_key(all_items, q_name, data_context):
    """方式Cのキャッシュキー。回答内容・設問名・分析データの特徴が同じなら同一キーになる（回答順序は無視）"""
    texts = sorted(x['text'] for x in all_items)
    src = q_name + '␟' + data_context + '␟' + '␞'.join(texts)
    return hashlib.sha256(src.encode('utf-8')).hexdigest()


# Stage1主題抽出の進捗をディスクに保存するディレクトリ。
# st.session_stateはアプリ再起動・セッション切断で消えてしまうため、
# 大量データ（全件精査など）でバッチ数が多い場合に途中経過を失わないよう、
# ディスクにも進捗を残し、次回同じデータで実行した際に未処理バッチから再開できるようにする。
STAGE1_CACHE_DIR = APP_DIR / '.codebook_cache'


def _stage1_cache_path(cache_key):
    # cache_keyには「:」（ratioの区切り）が含まれ、Windowsのファイル名では
    # 特殊な意味を持つ（NTFSの代替データストリームとして扱われ、ファイルが
    # 実質空になる）ため、ファイル名にはcache_key自体ではなく再ハッシュした値を使う。
    safe_name = hashlib.sha256(cache_key.encode('utf-8')).hexdigest()
    return STAGE1_CACHE_DIR / f'{safe_name}.json'


def _load_stage1_checkpoint(cache_key):
    path = _stage1_cache_path(cache_key)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return None


def _save_stage1_checkpoint(cache_key, data):
    STAGE1_CACHE_DIR.mkdir(exist_ok=True)
    _stage1_cache_path(cache_key).write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')


def _clear_stage1_checkpoint(cache_key):
    path = _stage1_cache_path(cache_key)
    if path.exists():
        path.unlink()


def _cleanup_old_stage1_checkpoints(max_age_hours=24):
    """古い（放棄された）チェックポイントファイルを掃除する。実行のたびに軽く呼ぶだけの簡易処理。"""
    if not STAGE1_CACHE_DIR.exists():
        return
    cutoff = time.time() - max_age_hours * 3600
    for f in STAGE1_CACHE_DIR.glob('*.json'):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
        except OSError:
            pass


def _build_codebook_c_stage1(client, stage1_items, max_codes, q_name, data_context, progress_bar, status_text,
                              cache_key):
    """
    方式C Stage1：ランダム抽出した一部から主題抽出→統合し初期コードブックを作る。
    大量データ（全件精査など）では主題抽出のバッチ数が多く1回の実行時間が長くなるため、
    バッチ完了ごとに進捗をディスクへ保存する。アプリの再起動やセッション切断で処理が
    中断しても、同じデータで再実行すれば未処理バッチから再開でき、最初からやり直しにならない。
    """
    _cleanup_old_stage1_checkpoints()
    batches = [stage1_items[i:i+50] for i in range(0, len(stage1_items), 50)]

    checkpoint = _load_stage1_checkpoint(cache_key)
    if checkpoint and checkpoint.get('total_batches') == len(batches):
        topics_all         = checkpoint.get('topics_all', [])
        completed_batches  = checkpoint.get('completed_batches', 0)
    else:
        topics_all        = []
        completed_batches = 0

    if completed_batches > 0:
        status_text.markdown(
            f'**Step 1/3** Stage1: 前回の続きから再開します（{completed_batches}/{len(batches)}バッチ済み）...'
        )

    for bi in range(completed_batches, len(batches)):
        status_text.markdown(f'**Step 1/3** Stage1（{len(stage1_items)}件）: 主題抽出中... {bi+1}/{len(batches)}バッチ')
        topics_all.extend(llm_extract_topics(client, batches[bi], q_name, data_context))
        progress_bar.progress(min(0.05 + 0.08 * ((bi+1) / len(batches)), 0.13))
        _save_stage1_checkpoint(cache_key, {
            'total_batches':     len(batches),
            'completed_batches': bi + 1,
            'topics_all':        topics_all,
        })

    topics_dedup = list(dict.fromkeys(topics_all))

    # 主題数が多いと統合コールが一括で処理しきれず失敗しやすいため、
    # チャンクごとにLLMで類似表現をまとめ、段階的に件数を減らしてから統合に渡す
    REDUCE_THRESHOLD = 300
    REDUCE_CHUNK = 200
    reduce_round = 1
    while len(topics_dedup) > REDUCE_THRESHOLD:
        status_text.markdown(f'**Step 1/3** Stage1: 主題リストを整理中...（{len(topics_dedup)}件、{reduce_round}回目）')
        chunks  = [topics_dedup[i:i+REDUCE_CHUNK] for i in range(0, len(topics_dedup), REDUCE_CHUNK)]
        reduced = []
        for chunk in chunks:
            reduced.extend(llm_reduce_topics(client, chunk, q_name, data_context))
        new_dedup = list(dict.fromkeys(reduced))
        if len(new_dedup) >= len(topics_dedup):
            break  # これ以上減らない場合は打ち切り、現状のリストで統合に進む
        topics_dedup = new_dedup
        reduce_round += 1

    status_text.markdown(
        f'**Step 1/3** Stage1: 初期コードブックを確定中...（{len(topics_all)}件 → 整理後{len(topics_dedup)}件）'
    )
    codebook = llm_consolidate_topics(client, topics_dedup, max_codes, q_name, data_context)
    progress_bar.progress(0.15)
    if codebook:
        # 成功した場合のみチェックポイントを消す。失敗時は残しておき、
        # 再実行時に抽出済みの主題リストから（再抽出せず）統合からやり直せるようにする。
        _clear_stage1_checkpoint(cache_key)
    return codebook


def _build_codebook_c(client, all_items, max_codes, q_name, data_context, progress_bar, status_text, ratio=0.25):
    """
    方式C1/C2/C3共通処理：全件精査
    Stage1: 全体のratio割合をランダム抽出して主題抽出→統合し初期コードブックを作成
    Stage2: 残り（1-ratio）を差分検出（ratio=1.0の場合は残りがないため実施しない）
    ステージ完了ごとの結果はセッション内でキャッシュし、後段の失敗時に前段からの
    やり直しを避ける（同一データ・設問名・分析データの特徴・ratioの場合のみ再利用）。
    さらにStage1内部（主題抽出）はバッチごとにディスクへも進捗を保存しており、
    大量データでStage1の途中にアプリが再起動・セッション切断しても再開できる（詳細は
    `_build_codebook_c_stage1`参照）。
    """
    cache_key   = _topics_cache_key(all_items, q_name, data_context) + f':{ratio}'
    stage_cache = st.session_state.setdefault('stage_codebook_cache', {})
    cached      = stage_cache.get(cache_key)
    stage_done  = cached['stage'] if cached else 0
    codebook    = cached['codebook'] if cached else None

    n1        = max(int(len(all_items) * ratio), 1)
    remaining = all_items[n1:]

    if stage_done >= 1:
        status_text.markdown('**Step 1/3** Stage1: キャッシュ済みの初期コードブックを再利用...')
        progress_bar.progress(0.15 if remaining else 0.30)
    else:
        codebook = _build_codebook_c_stage1(
            client, all_items[:n1], max_codes, q_name, data_context, progress_bar, status_text, cache_key
        )
        if not codebook:
            return None
        stage_done = 1
        stage_cache[cache_key] = {'stage': 1, 'codebook': codebook}

    if not remaining:
        progress_bar.progress(0.30)
        return codebook

    if stage_done >= 2:
        status_text.markdown('**Step 1/3** Stage2: キャッシュ済みの結果を再利用...')
        progress_bar.progress(0.30)
    else:
        status_text.markdown(f'**Step 1/3** Stage2: 残り{len(remaining)}件を差分検出中...')
        codebook = _diff_detect_loop(
            client, codebook, remaining, max_codes, q_name, progress_bar, p_start=0.15, p_end=0.30
        )
        stage_cache[cache_key] = {'stage': 2, 'codebook': codebook}

    return codebook


# ══════════════════════════════════════════════════
# メイン処理
# ══════════════════════════════════════════════════

def _generate_codebook_step(client, all_items, max_codes, q_name, data_context, progress_bar, status_text,
                             codebook_mode, existing_codebook=None):
    """Step1相当：指定方式でコードブックを生成（または既存コードブックを使用）して返す"""
    status_text.markdown('**コードブックを生成中...**')
    progress_bar.progress(0.05)

    if codebook_mode == 'EXISTING':
        status_text.markdown('**アップロードされたコードブックを使用します**')
        codebook = existing_codebook
        progress_bar.progress(0.30)
    elif codebook_mode == 'B':
        codebook = _build_codebook_b(client, all_items, max_codes, q_name, data_context, progress_bar, status_text)
    elif codebook_mode in CODEBOOK_MODE_C_RATIO:
        ratio = CODEBOOK_MODE_C_RATIO[codebook_mode]
        codebook = _build_codebook_c(client, all_items, max_codes, q_name, data_context, progress_bar, status_text, ratio)
    else:
        codebook = _build_codebook_a(client, all_items, max_codes, q_name, data_context, progress_bar)

    return codebook


CODING_SCOPE_OPTIONS = ['コーディングしない（策定のみ）', '100件でコーディング', '200件でコーディング', '全件コーディング']
CODING_SCOPE_SIZES   = {
    'コーディングしない（策定のみ）': 0,
    '100件でコーディング':          100,
    '200件でコーディング':          200,
    '全件コーディング':              None,  # Noneは実行時に全件数へ解決
}


def _code_items(client, q_name, codes, items, progress_bar, status_text, p_start=0.30, p_end=0.90,
                 coding_model=CODING_MODEL, enabled_risks=None):
    """itemsをコーディングし、結果リストを返す（集計は行わない）"""
    if not items:
        return []
    BATCH   = 15
    results = []
    batches = [items[i:i+BATCH] for i in range(0, len(items), BATCH)]
    for bi, batch in enumerate(batches):
        res = llm_code_batch(client, batch, codes, q_name, model=coding_model, enabled_risks=enabled_risks)
        results.extend(res)
        pct = p_start + (p_end - p_start) * ((bi+1) / len(batches))
        progress_bar.progress(min(pct, p_end))
        status_text.markdown(f'**コーディング中...** {bi+1}/{len(batches)}バッチ')
        time.sleep(0.1)
    return results


def run_pipeline(api_key, q_name, texts, max_codes, progress_bar, status_text, data_context='',
                  codebook_mode='A', existing_codebook=None, sample_size=None, coding_model=CODING_MODEL,
                  enabled_risks=None):
    """
    コードブック策定（または既存コードブックの読み込み）を行い、指定件数だけコーディング・集計する。
    sample_size=0: コーディングしない（策定のみ）／ None: 全件 ／ それ以外: min(sample_size, 全件数)件
    coding_model・enabled_risksはこの分析（result）に紐づけて保存し、続きをコーディングする際も
    同じ設定を使う（分析の途中でモデルやリスクチェック対象が混在しないようにするため）。
    """
    enabled_risks = enabled_risks or []
    reset_token_usage()
    client = make_client('Anthropic', api_key)
    all_items = [{'id': f'NO{i+1:03d}', 'text': t} for i, t in enumerate(texts)]
    random.shuffle(all_items)

    codebook = _generate_codebook_step(
        client, all_items, max_codes, q_name, data_context, progress_bar, status_text,
        codebook_mode, existing_codebook
    )
    if not codebook:
        reason = get_last_error() or '原因不明（AIから有効なコードブック構造が返されませんでした）'
        st.error(
            'コードブック生成に失敗しました。再度お試しください。'
            + f'\n\n詳細: {reason}'
        )
        return None

    codes = [
        {**c, 'cat_id': cat['cat_id'], 'cat_name': cat['cat_name']}
        for cat in codebook.get('categories', [])
        for c in cat.get('codes', [])
    ]
    status_text.markdown(f'**コードブック完成：{len(codes)}コード ✓**')
    progress_bar.progress(0.30)

    total_items = len(all_items)
    n = total_items if sample_size is None else min(sample_size, total_items)

    results = _code_items(client, q_name, codes, all_items[:n], progress_bar, status_text,
                           coding_model=coding_model, enabled_risks=enabled_risks)

    status_text.markdown('**集計中...**')
    gt, sent_counts, unassigned, risk_counts = aggregate_results(codes, results, n, risk_keys=enabled_risks)
    progress_bar.progress(1.0)
    status_text.markdown('**✅ 完了！**')

    usage = get_token_usage()
    return {
        'codebook':      codebook,
        'codes':         codes,
        'items':         all_items,
        'results':       results,
        'coded_count':   n,
        'total_items':   total_items,
        'gt':            gt,
        'sent':          sent_counts,
        'unassigned':    unassigned,
        'risk_counts':   risk_counts,
        'enabled_risks': enabled_risks,
        'q_name':        q_name,
        'usage':         usage,
        'coding_model':  coding_model,
    }


def continue_coding(api_key, q_name, codes, all_items, coded_count, add_size, progress_bar, status_text,
                     prior_results=None, prior_usage=None, coding_model=CODING_MODEL, enabled_risks=None):
    """既存のコーディング結果に続けて、未コーディング分から追加でadd_size件をコーディングする"""
    enabled_risks = enabled_risks or []
    reset_token_usage()
    client = make_client('Anthropic', api_key)

    total_items = len(all_items)
    new_count   = min(coded_count + add_size, total_items)
    new_items   = all_items[coded_count:new_count]

    new_results = _code_items(client, q_name, codes, new_items, progress_bar, status_text, 0.05, 0.90,
                               coding_model=coding_model, enabled_risks=enabled_risks)

    status_text.markdown('**集計中...**')
    results = (prior_results or []) + new_results
    gt, sent_counts, unassigned, risk_counts = aggregate_results(codes, results, new_count, risk_keys=enabled_risks)
    progress_bar.progress(1.0)
    status_text.markdown('**✅ 完了！**')

    usage = get_token_usage()
    if prior_usage:
        usage = {
            'input':          usage.get('input', 0)          + prior_usage.get('input', 0),
            'output':         usage.get('output', 0)         + prior_usage.get('output', 0),
            'cache_read':     usage.get('cache_read', 0)     + prior_usage.get('cache_read', 0),
            'cache_creation': usage.get('cache_creation', 0) + prior_usage.get('cache_creation', 0),
            'cost_jpy':       usage.get('cost_jpy', 0)       + prior_usage.get('cost_jpy', 0),
        }

    return {
        'results':     results,
        'coded_count': new_count,
        'gt':          gt,
        'sent':        sent_counts,
        'unassigned':  unassigned,
        'risk_counts': risk_counts,
        'usage':       usage,
    }
# ══════════════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════════════

st.markdown('<p class="main-title">👻 アフターコーディング支援ツール</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">アップロードした自由文回答テキストを自動でコーディングし集計します</p>', unsafe_allow_html=True)

# ── サイドバー：設定 ──────────────────────────────
with st.sidebar:
    st.markdown('# After Coder')
    st.markdown('*by Marketing Junction*')
    st.caption(f'👤 ログイン中: {st.session_state.username}')
    st.divider()
    st.header('⚙️ 設定')
    api_key = st.text_input(
        'APIキー',
        type='password',
        placeholder='sk-ant-...',
        help='Anthropic ConsoleでAPIキーを取得してください'
    )
    q_name = st.text_input(
        '設問名',
        placeholder='例：サービスへのご意見・ご感想',
        help='分析する自由回答設問の名前を入力してください'
    )
    data_context = st.text_area(
        label='分析データの特徴',
        placeholder='例：青森県の小中高校教員を対象とした教育改革に関するアンケートです。回答者は教員・校長・教頭・学校職員です。',
        height=120,
        help='コーディング作業時にAIが参考に使用します。調査内容や目的、調査対象者、質問内容など、データの特徴について記入してください。（記入しないでも作動はしますが記入した方が精度が上がります）'
    )
    max_codes = st.slider(
        'コード数の上限',
        min_value=10, max_value=49, value=30, step=1,
        help='生成するコードの最大数（推奨：20〜35）'
    )

    st.markdown('**📐 コードブック策定方式**')
    codebook_mode_label = st.selectbox(
        'コードブック策定方式',
        CODEBOOK_MODE_LABELS,
        index=1,
        label_visibility='collapsed',
    )
    codebook_mode = CODEBOOK_MODE_CODE[codebook_mode_label]

    n_texts = st.session_state.texts_count
    if n_texts >= 2000 and codebook_mode in ('A', 'B'):
        st.info(f'📊 {n_texts}件のデータには方式C1〜C3を推奨します')
    elif n_texts >= 500 and codebook_mode == 'A':
        st.info(f'📊 {n_texts}件のデータには方式B以上を推奨します')

    existing_codebook_data = None
    if codebook_mode == 'EXISTING':
        existing_file = st.file_uploader(
            'コードブックファイル（CSV または JSON、過去バージョンも可）',
            type=['csv', 'json'],
            help='「コードブック」または「コード一覧集計」の表右上のツールバーからダウンロードしたCSV、'
                 'または以前保存したJSONファイルを指定してください（過去のバージョンをアップロードして編集することもできます）'
        )
        if existing_file:
            try:
                if existing_file.name.lower().endswith('.csv'):
                    existing_codebook_data = parse_codebook_csv(existing_file.read())
                else:
                    existing_codebook_data = json.loads(existing_file.read().decode('utf-8'))
                n_loaded_codes = sum(
                    len(cat.get('codes', [])) for cat in existing_codebook_data.get('categories', [])
                )
                st.success(f'✅ コードブックを読み込みました（{n_loaded_codes}コード）')
            except Exception as e:
                st.error(f'ファイルの読み込みに失敗しました: {e}')
        st.caption('※ 既存のコードブックを使用する場合、「コード数の上限」は適用されません')

    st.markdown('**🧮 コーディング範囲**')
    coding_scope_label = st.selectbox(
        'コーディング範囲',
        CODING_SCOPE_OPTIONS,
        index=1,
        label_visibility='collapsed',
        help='コードブック策定（または既存コードブックの読み込み）の後、どこまでコーディングするかを選びます。'
             'まず少量でコーディングし、結果を見ながらコードブックを編集し、'
             '納得できたら結果画面の「続きをコーディングする」や「全件コーディング」で範囲を広げる使い方を想定しています。'
    )
    coding_sample_size = CODING_SCOPE_SIZES[coding_scope_label]

    coding_model_label = st.selectbox(
        'コーディングモデル',
        list(CODING_MODEL_OPTIONS.keys()),
        index=0,
        help='コーディング（分類作業）に使うモデルを選びます。Haiku 4.5はSonnet 4.6の約1/3の単価です。'
             '精度・価格を比較したい場合は、同じコードブックのまま方式を変えて2回実行し、'
             '「📜 作業履歴」で結果とコストを見比べてください（1回の分析中でモデルが混在することはありません）。'
    )
    coding_model = CODING_MODEL_OPTIONS[coding_model_label]

    st.markdown('**⚠️ リスクチェック（該当有無を検知）**')
    enabled_risks = []
    for opt in RISK_CHECK_OPTIONS:
        if st.checkbox(opt['label'], help=opt['hint'], key=f"risk_{opt['key']}"):
            enabled_risks.append(opt['key'])

    st.divider()
    st.markdown('**📖 使い方**')
    st.markdown('''
1. APIキーと設問名を入力
2. 分析データの特徴を記入（任意）
3. テキストファイルをアップロード
4. 「分析開始」をクリック
5. 結果を確認してExcelをダウンロード
    ''')

    st.divider()
    st.markdown('**📜 作業履歴**')
    if st.session_state.history:
        for h in reversed(st.session_state.history[-10:]):
            ts    = h['timestamp'].strftime('%m/%d %H:%M')
            label = h['q_name'] if len(h['q_name']) <= 18 else h['q_name'][:18] + '…'
            active = h['id'] == st.session_state.active_history_id
            if st.button(
                f"{'▶ ' if active else ''}{ts}　{label}",
                key=f"hist_btn_{h['id']}",
                width='stretch',
                type='primary' if active else 'secondary',
            ):
                st.session_state.active_history_id = h['id']
                st.rerun()
    else:
        st.caption('まだ分析履歴がありません')

    st.divider()
    if st.button('🚪 ログアウト', width='stretch'):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ── メインエリア ──────────────────────────────────
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown('### ⬆️ データアップロード')

    input_method = st.radio(
        '入力方法を選択',
        ['📄 ファイルをアップロード', '📋 テキストを直接貼り付け'],
        horizontal=True
    )

    texts = []

    if input_method == '📄 ファイルをアップロード':
        uploaded = st.file_uploader(
            'テキストファイルを選択（1行1回答）',
            type=['txt'],
            help='UTF-8形式のテキストファイル。1行に1つの回答を記入してください。'
        )
        if uploaded:
            content = uploaded.read().decode('utf-8', errors='ignore')
            texts   = [line.strip() for line in content.splitlines() if line.strip()]
            st.success(f'✅ {len(texts)}件の回答を読み込みました')
            with st.expander('回答プレビュー（先頭5件）'):
                for i, t in enumerate(texts[:5], 1):
                    st.markdown(f'**{i}.** {t}')

    else:
        pasted = st.text_area(
            'テキストを貼り付け（1行1回答）',
            height=200,
            placeholder='例：\n対応が丁寧で安心できた\n待ち時間が長かった\n先生の説明がわかりやすかった',
            help='1行に1つの回答を入力してください。空行は自動的に除外されます。'
        )
        if pasted:
            texts = [line.strip() for line in pasted.splitlines() if line.strip()]
            st.success(f'✅ {len(texts)}件の回答を読み込みました')
            with st.expander('回答プレビュー（先頭5件）'):
                for i, t in enumerate(texts[:5], 1):
                    st.markdown(f'**{i}.** {t}')

if st.session_state.texts_count != len(texts):
    st.session_state.texts_count = len(texts)
    st.rerun()

with col2:
    st.markdown('### ✅ 分析設定の確認')
    if texts and q_name and api_key:
        st.markdown(f'**設問名：** {q_name}')
        st.markdown(f'**回答数：** {len(texts)}件')
        st.markdown(f'**コード上限：** {max_codes}個')
        st.markdown(f'**策定方式：** {codebook_mode_label}')
        st.markdown(f'**コーディングモデル：** {coding_model_label}')
        risk_labels = [o['label'] for o in RISK_CHECK_OPTIONS if o['key'] in enabled_risks]
        st.markdown(f'**リスクチェック：** {"、".join(risk_labels) if risk_labels else "なし"}')
        st.markdown(f'**APIキー：** {"設定済み ✅" if api_key else "未設定"}')
        est_min = max(3, len(texts) // 100 * 2)
        st.info(f'⏱️ 処理時間の目安：{est_min}〜{est_min*2}分')
    else:
        st.info('左サイドバーで設定を入力し、ファイルをアップロードしてください。')
st.divider()

# ── 分析実行 ──────────────────────────────────────
mode_ready   = codebook_mode != 'EXISTING' or existing_codebook_data is not None
button_label = '📐 コードブック生成開始' if coding_sample_size == 0 else '🚀 分析開始'

if st.button(button_label, type='primary', width='stretch',
             disabled=not (texts and q_name and api_key and mode_ready)):

    progress_bar = st.progress(0)
    status_text  = st.empty()

    with st.spinner('処理中...'):
        result = run_pipeline(
            api_key, q_name, texts, max_codes,
            progress_bar, status_text, data_context, codebook_mode, existing_codebook_data,
            coding_sample_size, coding_model, enabled_risks
        )

    if result:
        st.session_state.history_counter += 1
        hist_entry = {
            'id':        st.session_state.history_counter,
            'q_name':    q_name,
            'timestamp': datetime.now(),
            'result':    result,
        }
        st.session_state.history.append(hist_entry)
        st.session_state.history = st.session_state.history[-10:]
        st.session_state.active_history_id = hist_entry['id']

active_result = None
active_q_name = None
for h in st.session_state.history:
    if h['id'] == st.session_state.active_history_id:
        active_result = h['result']
        active_q_name = h['q_name']
        break

if active_result:
    result = active_result
    q_name = active_q_name

    coded_count = result.get('coded_count', 0)
    total_items = result.get('total_items', 0)

    if coded_count == 0:
        st.success('✅ コードブックの生成が完了しました！')
    elif coded_count < total_items:
        st.success(f'✅ {coded_count}/{total_items}件をコーディングしました！')
    else:
        st.success('✅ 分析が完了しました！')

    if coded_count > 0:
        used_model = result.get('coding_model', CODING_MODEL)
        used_label = next((k for k, v in CODING_MODEL_OPTIONS.items() if v == used_model), used_model)
        st.caption(f'🧮 この分析のコーディングモデル：{used_label}（作業履歴から他の分析と比較できます）')

    # コスト表示（共通。プロンプトキャッシュの読み込み/書き込み分は呼び出し時点のモデル単価で
    # 都度計算し累積している＝コードブック策定と方式が混在していても正確な合計になる）
    usage      = result.get('usage', {})
    inp        = usage.get('input', 0)
    out        = usage.get('output', 0)
    cache_read = usage.get('cache_read', 0)
    cache_new  = usage.get('cache_creation', 0)
    cost       = usage.get('cost_jpy', 0.0)
    with st.expander('💰 API使用コスト（参考）'):
        c1, c2, c3 = st.columns(3)
        c1.metric('入力トークン', f'{inp:,}')
        c2.metric('出力トークン', f'{out:,}')
        c3.metric('推定コスト', f'約 ¥{cost:.0f}')
        if cache_read or cache_new:
            c4, c5 = st.columns(2)
            c4.metric('キャッシュ読込（約1/10単価）', f'{cache_read:,}')
            c5.metric('キャッシュ書込（約1.25倍単価）', f'{cache_new:,}')
        st.caption(
            '※ 1USD=150円換算。コードブック策定はSonnet、コーディングはHaikuの料金に基づく概算です。'
            'コーディングはコードブックをプロンプトキャッシュしており、2回目以降のバッチはキャッシュ読込分が割安になります。'
        )

    # ── コードブック（構造のみ。折りたたまず常時表示、編集直後もその場で最新反映） ──
    st.markdown('#### 📐 コードブック')
    render_codebook_structure(result['codebook'], key='codebook_current')

    # ── 編集指示の入力欄（コードブックの直下に固定） ──────────────
    st.caption(
        '⚠️ 編集すると現在の内容が置き換わります。コーディング済みの結果には自動反映されません'
        '（反映するには下の「現在のコードブックでコーディングする」で再度コーディングしてください）。'
        '保存しておきたい場合は、上の表の右上ツールバーから先にCSVをダウンロードしてください。'
    )
    with st.form('edit_instruction_form', clear_on_submit=True):
        instruction = st.text_input(
            '編集の指示を入力', label_visibility='collapsed',
            placeholder='例：AとBのコードを統合して／「対応の速さ」というコードを追加して定義は〜'
        )
        submitted = st.form_submit_button('編集案を作成する', width='stretch')

    if submitted and instruction:
        with st.spinner('編集案を作成中...'):
            client   = make_client('Anthropic', api_key)
            proposed = llm_edit_codebook(client, result['codebook'], instruction, result.get('q_name', q_name))
        if proposed:
            result['pending_edit'] = {'instruction': instruction, 'codebook': proposed}
            for h in st.session_state.history:
                if h['id'] == st.session_state.active_history_id:
                    h['result'] = result
                    break
            st.rerun()
        else:
            reason = get_last_error() or '原因不明（AIから有効なコードブック構造が返されませんでした）'
            st.error(f'編集案の作成に失敗しました。再度お試しください。\n\n詳細: {reason}')

    # ── 編集案のプレビュー（確定・キャンセルの2段階確認） ──────────
    pending_edit = result.get('pending_edit')
    if pending_edit:
        st.info(f"📝 編集案：「{pending_edit['instruction']}」（内容を確認して確定してください）")
        render_codebook_structure(pending_edit['codebook'], key='codebook_pending')
        pc1, pc2 = st.columns(2)
        with pc1:
            if st.button('✅ この内容で確定する', type='primary', width='stretch'):
                edit_log = result.setdefault(
                    'edit_log', [{'instruction': '（初期状態）', 'codebook': result['codebook']}]
                )
                edit_log.append({'instruction': pending_edit['instruction'], 'codebook': pending_edit['codebook']})
                result['codebook'] = pending_edit['codebook']
                result['codes'] = [
                    {**c, 'cat_id': cat['cat_id'], 'cat_name': cat['cat_name']}
                    for cat in pending_edit['codebook'].get('categories', [])
                    for c in cat.get('codes', [])
                ]
                del result['pending_edit']
                for h in st.session_state.history:
                    if h['id'] == st.session_state.active_history_id:
                        h['result'] = result
                        break
                st.rerun()
        with pc2:
            if st.button('❌ キャンセル', width='stretch'):
                del result['pending_edit']
                for h in st.session_state.history:
                    if h['id'] == st.session_state.active_history_id:
                        h['result'] = result
                        break
                st.rerun()

    # ── 編集履歴 ──────────────────────────────────────────
    edit_log = result.setdefault('edit_log', [{'instruction': '（初期状態）', 'codebook': result['codebook']}])
    if len(edit_log) > 1:
        with st.expander(f'📜 編集履歴（{len(edit_log)}バージョン）'):
            for i, entry in enumerate(edit_log):
                n_codes = sum(len(c.get('codes', [])) for c in entry['codebook'].get('categories', []))
                hc1, hc2 = st.columns([4, 1])
                hc1.markdown(f"**v{i}** {entry['instruction']}（コード{n_codes}件）")
                if i != len(edit_log) - 1:
                    if hc2.button('このバージョンに戻す', key=f'revert_edit_{i}'):
                        edit_log.append({'instruction': f'v{i}のバージョンに戻す', 'codebook': entry['codebook']})
                        result['codebook'] = entry['codebook']
                        result['codes'] = [
                            {**c, 'cat_id': cat['cat_id'], 'cat_name': cat['cat_name']}
                            for cat in entry['codebook'].get('categories', [])
                            for c in cat.get('codes', [])
                        ]
                        for h in st.session_state.history:
                            if h['id'] == st.session_state.active_history_id:
                                h['result'] = result
                                break
                        st.rerun()

    st.divider()

    # ── 現在のコードブックでコーディングする（1ボタン＋全件やり直しボタン） ──
    if pending_edit:
        st.caption('※ 編集案を確定またはキャンセルしてからコーディングしてください。')
    else:
        if coded_count < total_items:
            remaining = total_items - coded_count
            if st.button(f'▶ 現在のコードブックでコーディングする（残り{remaining}件）', type='primary', width='stretch'):
                progress_bar2 = st.progress(0)
                status_text2  = st.empty()
                with st.spinner('コーディング中...'):
                    continued = continue_coding(
                        api_key, result.get('q_name', q_name), result['codes'], result['items'],
                        coded_count, remaining, progress_bar2, status_text2,
                        prior_results=result.get('results', []), prior_usage=result.get('usage'),
                        coding_model=result.get('coding_model', CODING_MODEL),
                        enabled_risks=result.get('enabled_risks', []),
                    )
                result.update(continued)
                for h in st.session_state.history:
                    if h['id'] == st.session_state.active_history_id:
                        h['result'] = result
                        break
                st.rerun()
        if coded_count > 0:
            if st.button(f'🔁 現在のコードブックで全件コーディングし直す（全{total_items}件）', width='stretch'):
                progress_bar3 = st.progress(0)
                status_text3  = st.empty()
                with st.spinner('コーディング中...'):
                    recoded = continue_coding(
                        api_key, result.get('q_name', q_name), result['codes'], result['items'],
                        0, total_items, progress_bar3, status_text3,
                        prior_results=None, prior_usage=result.get('usage'),
                        coding_model=result.get('coding_model', CODING_MODEL),
                        enabled_risks=result.get('enabled_risks', []),
                    )
                result.update(recoded)
                for h in st.session_state.history:
                    if h['id'] == st.session_state.active_history_id:
                        h['result'] = result
                        break
                st.rerun()
            st.caption('※ コードの統合・削除などの編集をコーディング済みの回答すべてに反映したい場合は、こちらで最初からやり直してください。')
    st.divider()

    # ── コーディング結果（1件以上コーディング済みの場合のみ表示） ──────
    if coded_count > 0:
        st.subheader('📊 コーディング結果')

        sent         = result['sent']
        unassigned   = result.get('unassigned', 0)
        risk_counts  = result.get('risk_counts', {})
        result_risks = result.get('enabled_risks', [])
        def _pct(cnt):
            return f'{cnt/coded_count*100:.1f}%' if coded_count else '0.0%'

        risk_icons = {'claim': '📣', 'personal': '🪪', 'address': '📍', 'org': '🏢', 'danger': '🚨'}
        risk_html = ''.join(
            f"  <div>{risk_icons.get(o['key'], '⚠️')} <b>{o['label']}</b>　"
            f"{risk_counts.get(o['key'], 0)}件（{_pct(risk_counts.get(o['key'], 0))}）</div>\n"
            for o in RISK_CHECK_OPTIONS if o['key'] in result_risks
        )
        with st.expander('📌 特記情報', expanded=True):
            st.markdown(f"""
<div style="display:flex; flex-wrap:nowrap; overflow-x:auto; gap:32px; padding:4px 0;">
  <div>😊 <b>ポジティブ</b>　{sent['positive']}件（{_pct(sent['positive'])}）</div>
  <div>😞 <b>ネガティブ</b>　{sent['negative']}件（{_pct(sent['negative'])}）</div>
  <div>😐 <b>ニュートラル</b>　{sent['neutral']}件（{_pct(sent['neutral'])}）</div>
  <div>➖ <b>非該当（コードなし）</b>　{unassigned}件（{_pct(unassigned)}）</div>
{risk_html}</div>
""", unsafe_allow_html=True)

        st.divider()

        gt = result['gt']
        import pandas as pd

        with st.expander('カテゴリ別集計を表示'):
            cat_summary = {}
            for item in gt:
                cid = item['cat_id']
                if cid not in cat_summary:
                    cat_summary[cid] = {'カテゴリID': cid, 'カテゴリ名': item['cat_name'], '件数': 0}
                cat_summary[cid]['件数'] += item['count']
            df_cat = pd.DataFrame(list(cat_summary.values()))
            df_cat['出現率(%)'] = (df_cat['件数'] / coded_count * 100).round(1)
            df_cat = df_cat.sort_values('件数', ascending=False).reset_index(drop=True)
            st.dataframe(df_cat, width='stretch', hide_index=True)

        gt_by_code = {g['code_id']: {'count': g['count'], 'pct': g['pct']} for g in gt}
        render_code_list_table(result['codebook'], gt_by_code=gt_by_code, key='code_list_table')

        st.divider()

        st.markdown('#### 📈 コード別GT集計')

        sort_mode = st.radio(
            '表示順',
            ['順A：カテゴリ出現率順→コード出現率順', '順B：コード出現率が多い順'],
            horizontal=True
        )

        cat_total = {}
        for item in gt:
            cat_total[item['cat_id']] = cat_total.get(item['cat_id'], 0) + item['count']

        if sort_mode == '順B：コード出現率が多い順':
            gt_sorted = sorted(gt, key=lambda x: x['count'], reverse=True)
        else:
            gt_sorted = sorted(gt, key=lambda x: (-cat_total.get(x['cat_id'], 0), -x['count']))

        import plotly.express as px
        df_plot = pd.DataFrame(gt_sorted)[['cat_name','code_name','count','pct']]
        df_plot.columns = ['カテゴリ','コード名','件数','出現率(%)']
        fig = px.bar(
            df_plot,
            x='コード名',
            y='出現率(%)',
            color='カテゴリ',
            category_orders={'コード名': df_plot['コード名'].tolist()},
            labels={'出現率(%)': '出現率(%)', 'コード名': ''},
        )
        fig.update_layout(
            xaxis_tickangle=-45,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
            height=500,
            margin=dict(b=120),
        )
        st.plotly_chart(fig, width='stretch')

        st.divider()

        partial_note = '' if coded_count >= total_items else f'（{coded_count}/{total_items}件分）'
        st.markdown(f'#### 💾 レポートのダウンロード{partial_note}')
        excel_bytes = create_excel(
            q_name, gt, sent, coded_count,
            result['results'], result['items'][:coded_count], result['codes'], unassigned,
            risk_counts=result.get('risk_counts', {}), enabled_risks=result.get('enabled_risks', []),
        )
        st.download_button(
            label='📥 Excelレポートをダウンロード',
            data=excel_bytes,
            file_name=f'AfterCoding_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            width='stretch',
        )