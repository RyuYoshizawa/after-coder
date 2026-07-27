"""
アフターコーディング支援ツール - デモ版
Streamlit Webアプリ
"""

import streamlit as st
import anthropic
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
from llm_client import call_llm, make_client, reset_token_usage, get_token_usage, calc_cost_jpy

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
    'user01': 'pass01',
    'user02': 'pass02',
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
# ══════════════════════════════════════════════════
# 集計・Excel出力関数
# ══════════════════════════════════════════════════

def aggregate_results(codes, results, total):
    """コード別GT集計とセンチメント集計"""
    code_counts = {c['code_id']: 0 for c in codes}
    sent_counts = {'positive': 0, 'negative': 0, 'neutral': 0}
    result_map  = {r['id']: r for r in results}

    for rid, res in result_map.items():
        sent = res.get('sentiment', 'neutral')
        if sent in sent_counts:
            sent_counts[sent] += 1
        for cid in res.get('codes', []):
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
    return gt, sent_counts


def create_excel(q_name, gt, sent_counts, total, results, items, codes):
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

    # ── センチメント集計（D5〜I6） ────────────────────────────────
    ws.cell(row=4, column=1, value='■ センチメント集計').font = SUB_FONT
    sent_order = [('positive','ポジティブ'), ('negative','ネガティブ'), ('neutral','ニュートラル')]
    for i, (sent, label) in enumerate(sent_order):
        cnt = sent_counts[sent]
        pct = cnt / total * 100 if total > 0 else 0
        hdr(5, 4+i*2, label)
        dat(5, 5+i*2, cnt)
        hdr(6, 4+i*2, '%')
        dat(6, 5+i*2, round(pct, 1))

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
    for i, (sent, label) in enumerate([('positive','ポジティブ'),('negative','ネガティブ'),('neutral','ニュートラル')]):
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
# メイン処理
# ══════════════════════════════════════════════════

def run_analysis(api_key, q_name, texts, max_codes, progress_bar, status_text, data_context=''):
    """コードブック策定→全件コーディング→集計を実行"""

    reset_token_usage()
    client = make_client('Anthropic', api_key)
    all_items = [{'id': f'NO{i+1:03d}', 'text': t} for i, t in enumerate(texts)]
    random.shuffle(all_items)
    total = len(all_items)

    # ── Step1: コードブック策定 ──────────────────────────────────
    status_text.markdown('**Step 1/3** コードブックを生成中...')
    progress_bar.progress(0.05)

    sample_1   = all_items[:30]
    codebook   = llm_generate_codebook(client, sample_1, max_codes, q_name, data_context)
    if not codebook:
        st.error('コードブック生成に失敗しました。再度お試しください。')
        return None

    # 差分検出（残りを20件ずつ）
    remaining = all_items[30:]
    round_no  = 2
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
        progress_bar.progress(min(0.05 + 0.25 * (round_no / max(len(all_items)//20, 1)), 0.30))
        round_no += 1

    # コードをフラットリストに
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
    gt, sent_counts = aggregate_results(codes, results, total)
    progress_bar.progress(1.0)
    status_text.markdown('**✅ 分析完了！**')

    usage = get_token_usage()
    return {
        'codebook': codebook,
        'codes':    codes,
        'results':  results,
        'gt':       gt,
        'sent':     sent_counts,
        'total':    total,
        'items':    all_items,
        'usage':    usage,
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

with col2:
    st.markdown('### ✅ 分析設定の確認')
    if texts and q_name and api_key:
        st.markdown(f'**設問名：** {q_name}')
        st.markdown(f'**回答数：** {len(texts)}件')
        st.markdown(f'**コード上限：** {max_codes}個')
        st.markdown(f'**APIキー：** {"設定済み ✅" if api_key else "未設定"}')
        est_min = max(3, len(texts) // 100 * 2)
        st.info(f'⏱️ 処理時間の目安：{est_min}〜{est_min*2}分')
    else:
        st.info('左サイドバーで設定を入力し、ファイルをアップロードしてください。')
st.divider()

# ── 分析実行 ──────────────────────────────────────
if st.button('🚀 分析開始', type='primary', use_container_width=True,
             disabled=not (texts and q_name and api_key)):

    progress_bar = st.progress(0)
    status_text  = st.empty()

    with st.spinner('分析中...'):
        result = run_analysis(
            api_key, q_name, texts, max_codes,
            progress_bar, status_text, data_context
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
    if True:
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

        st.divider()

        # ── 結果表示 ──────────────────────────────
        st.subheader('📊 分析結果')

        # センチメント集計
        st.markdown('#### 😊 センチメント集計')
        c1, c2, c3 = st.columns(3)
        total = result['total']
        sent  = result['sent']
        with c1:
            cnt = sent['positive']
            st.metric('ポジティブ', f'{cnt}件', f'{cnt/total*100:.1f}%')
        with c2:
            cnt = sent['negative']
            st.metric('ネガティブ', f'{cnt}件', f'{cnt/total*100:.1f}%')
        with c3:
            cnt = sent['neutral']
            st.metric('ニュートラル', f'{cnt}件', f'{cnt/total*100:.1f}%')

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
            result['results'], result['items'], result['codes']
        )
        st.download_button(
            label='📥 Excelレポートをダウンロード',
            data=excel_bytes,
            file_name=f'AfterCoding_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )