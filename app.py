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
from datetime import datetime
from llm_client import call_llm, make_client, reset_token_usage, get_token_usage, calc_cost_jpy, get_last_error

APP_DIR = Path(__file__).parent

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
- コードID形式: CAT01.../C0101...

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
    return call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')


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
    return call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')


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
- コードID形式: CAT01.../C0101..."""

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
    return call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')


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
    result = call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')
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
    result = call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')
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
- コードID形式: CAT01.../C0101..."""

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
    return call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')


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
cat_idは必ず既存カテゴリ一覧のCAT01形式を使うこと。"""

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
    result = call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')
    if result is None:
        return []
    return result.get('new_codes', [])


def llm_code_batch(client, items, codes, q_name):
    code_list = '\n'.join(
        f'{c["code_id"]}（{c["cat_name"]}）: {c["code_name"]} / {c["definition"][:25]}'
        for c in codes
    )
    items_text = '\n'.join(f'{x["id"]}: {x["text"]}' for x in items)
    prompt = f"""「{q_name}」の回答にコードブックに基づいてコーディングしてください。

【コードブック】
{code_list}

【ルール】
- 1回答に複数コード付与可
- 該当なしはcodesを空配列
- sentimentは positive/negative/neutral

【回答】
{items_text}"""

    schema = {
        'type': 'object',
        'properties': {
            'results': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'id':        {'type': 'string'},
                        'codes':     {'type': 'array', 'items': {'type': 'string'}},
                        'sentiment': {'type': 'string', 'enum': ['positive','negative','neutral']},
                    },
                    'required': ['id', 'codes', 'sentiment'],
                }
            }
        },
        'required': ['results'],
    }
    result = call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')
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
- コードID・カテゴリIDの形式は維持する（新規追加時はCAT.../C....形式で採番する）
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
    return call_llm(client, prompt, schema, 'Anthropic', 'claude-sonnet-4-6')
# ══════════════════════════════════════════════════
# 集計・Excel出力関数
# ══════════════════════════════════════════════════

def aggregate_results(codes, results, total):
    """コード別GT集計とセンチメント集計・非該当（どのコードも付与されなかった回答）件数を算出"""
    code_counts = {c['code_id']: 0 for c in codes}
    sent_counts = {'positive': 0, 'negative': 0, 'neutral': 0}
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
    return gt, sent_counts, unassigned


CODEBOOK_CSV_COLUMNS = ['カテゴリID', 'カテゴリ名', 'コードID', 'コード名', '定義', 'キーワード']


def render_codebook_table(codebook, expanded=False):
    """
    生成・使用したコードブックをカテゴリ→コード形式の表で折りたたみ表示する。
    表右上のツールバーからCSVダウンロードでき、そのCSVは「既存のコードブックを使用」で再読み込みできる。
    """
    import pandas as pd
    codes = [
        {**c, 'cat_id': cat['cat_id'], 'cat_name': cat['cat_name']}
        for cat in codebook.get('categories', [])
        for c in cat.get('codes', [])
    ]
    n_cats = len(codebook.get('categories', []))
    with st.expander(f'📐 コードブックを表示（カテゴリ{n_cats}／コード{len(codes)}）', expanded=expanded):
        df = pd.DataFrame(codes)
        show_cols = [c for c in ['cat_id', 'cat_name', 'code_id', 'code_name', 'definition', 'keywords'] if c in df.columns]
        df = df[show_cols].rename(columns={
            'cat_id': 'カテゴリID', 'cat_name': 'カテゴリ名', 'code_id': 'コードID',
            'code_name': 'コード名', 'definition': '定義', 'keywords': 'キーワード',
        })
        if 'キーワード' in df.columns:
            df['キーワード'] = df['キーワード'].apply(lambda kw: '; '.join(kw) if isinstance(kw, list) else (kw or ''))
        st.dataframe(df, width='stretch', hide_index=True)


def parse_codebook_csv(file_bytes):
    """
    render_codebook_tableの表からダウンロードしたCSV（カテゴリID,カテゴリ名,コードID,コード名,定義,キーワード）を
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


def create_excel(q_name, gt, sent_counts, total, results, items, codes, unassigned=0):
    """ローカル版仮集計シートと同じレイアウトでExcelを生成"""
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

    # カテゴリ別カラーパレット
    cat_colors = [
        'D9E2F3','FCE4D6','E2EFDA','FFF2CC','EDEDED',
        'F4CCCC','D0E4F5','EAD1DC','D9D2E9','CFE2F3'
    ]
    cat_ids   = list(dict.fromkeys(c['cat_id'] for c in codes))
    cat_color = {cid: cat_colors[i % len(cat_colors)] for i, cid in enumerate(cat_ids)}

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

    FIXED_N    = 3   # 回答ID・回答テキスト・センチメントの3列
    CODE_START = FIXED_N + 1  # D列からコード開始

    # ── センチメント集計（D5〜K6） ────────────────────────────────
    ws.cell(row=4, column=1, value='■ センチメント集計').font = SUB_FONT
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
        fill = PatternFill('solid',
                           start_color=cat_color.get(code['cat_id'], 'FFFFFF'),
                           end_color=cat_color.get(code['cat_id'], 'FFFFFF'))
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
            c.fill = fill
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

    for ci, code in enumerate(codes):
        col  = CODE_START + ci
        fill = PatternFill('solid',
                           start_color=cat_color.get(code['cat_id'], 'FFFFFF'),
                           end_color=cat_color.get(code['cat_id'], 'FFFFFF'))
        c = ws.cell(row=hdr_row, column=col, value=code['code_id'])
        c.font = Font(name='Meiryo UI', bold=True, size=10, color='000000')
        c.fill=fill; c.border=BORDER

    item_map = {item['id']: item['text'] for item in items}
    for ri, res in enumerate(results):
        r        = hdr_row + 1 + ri
        rid      = res.get('id', '')
        text     = item_map.get(rid, '')[:50]
        sent     = res.get('sentiment', '')
        assigned = res.get('codes', [])
        for ci, val in enumerate([rid, text, sent]):
            dat(r, 1+ci, val)
        for ci, code in enumerate(codes):
            col  = CODE_START + ci
            flag = 1 if code['code_id'] in assigned else 0
            fill = PatternFill('solid',
                               start_color=cat_color.get(code['cat_id'], 'FFFFFF'),
                               end_color=cat_color.get(code['cat_id'], 'FFFFFF'))
            c = ws.cell(row=r, column=col, value=flag if flag else '')
            c.font=DATA_FONT; c.border=BORDER
            if flag: c.fill = FLAG_FILL

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 12
    ws.freeze_panes = f'A{hdr_row+1}'

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

    # センチメント集計
    ws2.cell(row=4, column=1, value='■ センチメント集計').font = SUB_FONT
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

    # GT集計（行方向・出現率順）
    ws2.cell(row=8, column=1, value='■ コード別GT集計').font = SUB_FONT
    gt_headers = ['カテゴリID', 'カテゴリ名', 'コードID', 'コード名', '出現件数', '出現率(%)', '定義']
    for ci, h in enumerate(gt_headers):
        c = ws2.cell(row=9, column=ci+1, value=h)
        c.font=HDR_FONT; c.fill=HDR_FILL; c.border=BORDER

    # 出現率順にソート
    gt_sorted = sorted(gt, key=lambda x: x['count'], reverse=True)
    for ri, row_data in enumerate(gt_sorted):
        r    = 10 + ri
        pct  = row_data['pct']
        fill = PatternFill('solid',
                           start_color=cat_color.get(row_data['cat_id'], 'FFFFFF'),
                           end_color=cat_color.get(row_data['cat_id'], 'FFFFFF'))
        vals = [
            row_data['cat_id'],
            row_data['cat_name'],
            row_data['code_id'],
            row_data['code_name'],
            row_data['count'],
            pct,
            row_data['definition'],
        ]
        for ci, val in enumerate(vals):
            c = ws2.cell(row=r, column=ci+1, value=val)
            c.font=DATA_FONT; c.border=BORDER
            c.fill = fill
            if ci == 5:  # 出現率列
                if pct >= 20:
                    c.fill = HIGH_FILL
                elif pct <= 2:
                    c.fill = LOW_FILL
                    c.font = Font(name='Meiryo UI', size=10, color='FF0000', bold=True)

    # 列幅
    for ci, w in enumerate([10, 22, 10, 24, 10, 10, 40]):
        ws2.column_dimensions[get_column_letter(ci+1)].width = w
    ws2.freeze_panes = 'A10'

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


def _build_codebook_c_stage1(client, stage1_items, max_codes, q_name, data_context, progress_bar, status_text):
    """方式C Stage1：ランダム抽出した一部から主題抽出→統合し初期コードブックを作る"""
    batches    = [stage1_items[i:i+50] for i in range(0, len(stage1_items), 50)]
    topics_all = []
    for bi, batch in enumerate(batches):
        status_text.markdown(f'**Step 1/3** Stage1（{len(stage1_items)}件）: 主題抽出中... {bi+1}/{len(batches)}バッチ')
        topics_all.extend(llm_extract_topics(client, batch, q_name, data_context))
        progress_bar.progress(min(0.05 + 0.08 * ((bi+1) / len(batches)), 0.13))

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
    return codebook


def _build_codebook_c(client, all_items, max_codes, q_name, data_context, progress_bar, status_text, ratio=0.25):
    """
    方式C1/C2/C3共通処理：全件精査
    Stage1: 全体のratio割合をランダム抽出して主題抽出→統合し初期コードブックを作成
    Stage2: 残り（1-ratio）を差分検出（ratio=1.0の場合は残りがないため実施しない）
    途中のステージまでの結果はセッション内でキャッシュし、後段の失敗時に前段からの
    やり直しを避ける（同一データ・設問名・分析データの特徴・ratioの場合のみ再利用）。
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
            client, all_items[:n1], max_codes, q_name, data_context, progress_bar, status_text
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


def _code_and_aggregate(client, q_name, codebook, all_items, progress_bar, status_text):
    """Step2（全件コーディング）＋Step3（集計）の共通処理。トークン使用量の集計は呼び出し側で行う"""
    total = len(all_items)
    codes = [
        {**c, 'cat_id': cat['cat_id'], 'cat_name': cat['cat_name']}
        for cat in codebook.get('categories', [])
        for c in cat.get('codes', [])
    ]
    status_text.markdown(f'**Step 1/3** コードブック完成：{len(codes)}コード ✓')
    progress_bar.progress(0.30)

    # ── Step2: 全件コーディング ───────────────────────────────────
    status_text.markdown('**Step 2/3** 全件コーディング中...')
    BATCH   = 15
    results = []
    batches = [all_items[i:i+BATCH] for i in range(0, total, BATCH)]
    for bi, batch in enumerate(batches):
        res = llm_code_batch(client, batch, codes, q_name)
        results.extend(res)
        pct = 0.30 + 0.60 * ((bi+1) / len(batches))
        progress_bar.progress(min(pct, 0.90))
        status_text.markdown(f'**Step 2/3** コーディング中... {bi+1}/{len(batches)}バッチ')
        time.sleep(0.1)

    status_text.markdown(f'**Step 2/3** コーディング完了：{len(results)}件 ✓')
    progress_bar.progress(0.92)

    # ── Step3: 集計 ───────────────────────────────────────────────
    status_text.markdown('**Step 3/3** 集計中...')
    gt, sent_counts, unassigned = aggregate_results(codes, results, total)
    progress_bar.progress(1.0)
    status_text.markdown('**✅ 分析完了！**')

    return {
        'codebook':   codebook,
        'codes':      codes,
        'results':    results,
        'gt':         gt,
        'sent':       sent_counts,
        'unassigned': unassigned,
        'total':      total,
        'items':      all_items,
    }


def run_codebook_only(api_key, q_name, texts, max_codes, progress_bar, status_text, data_context='',
                       codebook_mode='A', existing_codebook=None):
    """コードブック策定のみを実行し、全件コーディング・集計はスキップする（策定方式の試行／既存コードブックの編集用）"""

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
    progress_bar.progress(1.0)
    status_text.markdown('**✅ コードブック生成完了！**')

    usage = get_token_usage()
    return {
        'codebook_only': True,
        'codebook':      codebook,
        'codes':         codes,
        'items':         all_items,
        'q_name':        q_name,
        'usage':         usage,
    }


def run_coding_from_codebook(api_key, q_name, codebook, all_items, progress_bar, status_text, prior_usage=None):
    """策定済み（生成済み・アップロード済み問わず）のコードブックを使って全件コーディング→集計のみ実行"""
    reset_token_usage()
    client = make_client('Anthropic', api_key)

    result = _code_and_aggregate(client, q_name, codebook, all_items, progress_bar, status_text)

    usage = get_token_usage()
    if prior_usage:
        usage = {
            'input':  usage.get('input', 0)  + prior_usage.get('input', 0),
            'output': usage.get('output', 0) + prior_usage.get('output', 0),
        }
    result['usage'] = usage
    return result


def run_analysis(api_key, q_name, texts, max_codes, progress_bar, status_text, data_context='',
                  codebook_mode='A', existing_codebook=None):
    """コードブック策定（または既存コードブックの再利用）→全件コーディング→集計を実行"""

    reset_token_usage()
    client = make_client('Anthropic', api_key)
    all_items = [{'id': f'NO{i+1:03d}', 'text': t} for i, t in enumerate(texts)]
    random.shuffle(all_items)

    # ── Step1: コードブック策定 ──────────────────────────────────
    status_text.markdown('**Step 1/3** コードブックを生成中...')
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

    result = _code_and_aggregate(client, q_name, codebook, all_items, progress_bar, status_text)
    result['usage'] = get_token_usage()
    return result
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
    codebook_mode_label = st.radio(
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
            help='「コードブックを表示」の表右上のツールバーからダウンロードしたCSV、'
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

    run_full_pipeline = st.checkbox(
        'コーディングまで全て実行する',
        value=False,
        help='デフォルトではコードブック策定（または既存コードブックの読み込み）のみを行い、'
             '内容を確認・編集してから「▶ コーディングを実行する」で続行できます。'
             'チェックすると、策定から全件コーディング・集計まで一気に実行します。'
    )

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
        st.markdown(f'**APIキー：** {"設定済み ✅" if api_key else "未設定"}')
        est_min = max(3, len(texts) // 100 * 2)
        st.info(f'⏱️ 処理時間の目安：{est_min}〜{est_min*2}分')
    else:
        st.info('左サイドバーで設定を入力し、ファイルをアップロードしてください。')
st.divider()

# ── 分析実行 ──────────────────────────────────────
mode_ready   = codebook_mode != 'EXISTING' or existing_codebook_data is not None
button_label = '🚀 分析開始' if run_full_pipeline else '📐 コードブック生成開始'

if st.button(button_label, type='primary', width='stretch',
             disabled=not (texts and q_name and api_key and mode_ready)):

    progress_bar = st.progress(0)
    status_text  = st.empty()

    with st.spinner('処理中...'):
        if not run_full_pipeline:
            result = run_codebook_only(
                api_key, q_name, texts, max_codes,
                progress_bar, status_text, data_context, codebook_mode, existing_codebook_data
            )
        else:
            result = run_analysis(
                api_key, q_name, texts, max_codes,
                progress_bar, status_text, data_context, codebook_mode, existing_codebook_data
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

    if result.get('codebook_only'):
        st.success('✅ コードブックの生成が完了しました！')

        usage = result.get('usage', {})
        inp   = usage.get('input', 0)
        out   = usage.get('output', 0)
        cost  = calc_cost_jpy(inp, out, 'claude-sonnet-4-6')
        with st.expander('💰 API使用コスト（参考）'):
            c1, c2, c3 = st.columns(3)
            c1.metric('入力トークン', f'{inp:,}')
            c2.metric('出力トークン', f'{out:,}')
            c3.metric('推定コスト', f'約 ¥{cost:.0f}')
            st.caption('※ 1USD=150円換算。claude-sonnet-4-6の料金に基づく概算です。')

        st.divider()

        render_codebook_table(result['codebook'], expanded=True)

        st.divider()
        st.markdown('#### 💬 コードブックを編集')
        st.caption(
            '⚠️ 編集すると現在の内容が置き換わります。保存しておきたい場合は、'
            '上の表の右上ツールバーから先にCSVをダウンロードしてください。'
        )

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

        for entry in edit_log[1:]:
            with st.chat_message('user'):
                st.write(entry['instruction'])
            with st.chat_message('assistant'):
                n_cats  = len(entry['codebook'].get('categories', []))
                n_codes = sum(len(c.get('codes', [])) for c in entry['codebook'].get('categories', []))
                st.write(f'更新しました（カテゴリ{n_cats}／コード{n_codes}）')

        instruction = st.chat_input(
            '編集の指示を入力（例：AとBのコードを統合して／「対応の速さ」というコードを追加して定義は〜）'
        )
        if instruction:
            with st.spinner('編集中...'):
                client  = make_client('Anthropic', api_key)
                updated = llm_edit_codebook(client, result['codebook'], instruction, result.get('q_name', q_name))
            if updated:
                edit_log.append({'instruction': instruction, 'codebook': updated})
                result['codebook'] = updated
                result['codes'] = [
                    {**c, 'cat_id': cat['cat_id'], 'cat_name': cat['cat_name']}
                    for cat in updated.get('categories', [])
                    for c in cat.get('codes', [])
                ]
                for h in st.session_state.history:
                    if h['id'] == st.session_state.active_history_id:
                        h['result'] = result
                        break
                st.rerun()
            else:
                reason = get_last_error() or '原因不明（AIから有効なコードブック構造が返されませんでした）'
                st.error(f'編集に失敗しました。再度お試しください。\n\n詳細: {reason}')

        st.divider()
        if st.button('▶ コーディングを実行する', type='primary', width='stretch'):
            progress_bar2 = st.progress(0)
            status_text2  = st.empty()
            with st.spinner('コーディング中...'):
                continued = run_coding_from_codebook(
                    api_key, result.get('q_name', q_name), result['codebook'], result['items'],
                    progress_bar2, status_text2, prior_usage=result.get('usage')
                )
            if continued:
                for h in st.session_state.history:
                    if h['id'] == st.session_state.active_history_id:
                        h['result'] = continued
                        break
                st.rerun()

    else:
        st.success('✅ 分析が完了しました！')

        # コスト表示
        usage = result.get('usage', {})
        inp   = usage.get('input', 0)
        out   = usage.get('output', 0)
        cost  = calc_cost_jpy(inp, out, 'claude-sonnet-4-6')
        with st.expander('💰 API使用コスト（参考）'):
            c1, c2, c3 = st.columns(3)
            c1.metric('入力トークン', f'{inp:,}')
            c2.metric('出力トークン', f'{out:,}')
            c3.metric('推定コスト', f'約 ¥{cost:.0f}')
            st.caption('※ 1USD=150円換算。claude-sonnet-4-6の料金に基づく概算です。')

        render_codebook_table(result['codebook'])

        st.divider()

        # ── 結果表示 ──────────────────────────────
        st.subheader('📊 分析結果')

        # センチメント集計
        st.markdown('#### 😊 センチメント集計')
        c1, c2, c3, c4 = st.columns(4)
        total      = result['total']
        sent       = result['sent']
        unassigned = result.get('unassigned', 0)
        with c1:
            cnt = sent['positive']
            st.metric('ポジティブ', f'{cnt}件', f'{cnt/total*100:.1f}%')
        with c2:
            cnt = sent['negative']
            st.metric('ネガティブ', f'{cnt}件', f'{cnt/total*100:.1f}%')
        with c3:
            cnt = sent['neutral']
            st.metric('ニュートラル', f'{cnt}件', f'{cnt/total*100:.1f}%')
        with c4:
            st.metric('非該当（コードなし）', f'{unassigned}件', f'{unassigned/total*100:.1f}%')

        st.divider()

        # GT集計
        st.markdown('#### 📈 コード別GT集計')
        gt = result['gt']
        import pandas as pd

        # 表示順の切り替え
        sort_mode = st.radio(
            '表示順',
            ['順A：カテゴリ出現率順→コード出現率順', '順B：コード出現率が多い順'],
            horizontal=True
        )

        # カテゴリ別出現数を集計
        cat_total = {}
        for item in gt:
            cat_total[item['cat_id']] = cat_total.get(item['cat_id'], 0) + item['count']

        if sort_mode == '順B：コード出現率が多い順':
            gt_sorted = sorted(gt, key=lambda x: x['count'], reverse=True)
        else:
            gt_sorted = sorted(gt, key=lambda x: (-cat_total.get(x['cat_id'], 0), -x['count']))

        # plotlyでカテゴリ別色分けグラフ
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
        st.plotly_chart(fig, use_container_width=True)

        # カテゴリ集計
        with st.expander('カテゴリ別集計を表示'):
            cat_summary = {}
            for item in gt_sorted:
                cid = item['cat_id']
                if cid not in cat_summary:
                    cat_summary[cid] = {'カテゴリID': cid, 'カテゴリ名': item['cat_name'], '件数': 0}
                cat_summary[cid]['件数'] += item['count']
            df_cat = pd.DataFrame(list(cat_summary.values()))
            df_cat['出現率(%)'] = (df_cat['件数'] / total * 100).round(1)
            df_cat = df_cat.sort_values('件数', ascending=False).reset_index(drop=True)
            st.dataframe(df_cat, use_container_width=True, hide_index=True)

        # 全コード一覧
        with st.expander('全コード一覧を表示'):
            df_full = pd.DataFrame(gt_sorted)[['cat_name','code_id','code_name','count','pct','definition']]
            df_full.columns = ['カテゴリ','コードID','コード名','件数','出現率(%)','定義']
            st.dataframe(df_full, use_container_width=True, hide_index=True)

        st.divider()

        # Excel ダウンロード
        st.markdown('#### 💾 レポートのダウンロード')
        excel_bytes = create_excel(
            q_name, gt, sent, total,
            result['results'], result['items'], result['codes'], unassigned
        )
        st.download_button(
            label='📥 Excelレポートをダウンロード',
            data=excel_bytes,
            file_name=f'AfterCoding_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            width='stretch',
        )