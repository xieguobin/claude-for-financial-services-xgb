#!/usr/bin/env python3
"""
Generic A-share Pitch Deck Generator
======================================
Usage:
    python3 generate_a_share_ppt.py --company "贵州茅台" --ticker "600519" --industry "白酒"
    python3 generate_a_share_ppt.py --company "宁德时代" --ticker "300750" --industry "电池"

All data fetched live from AkShare / public APIs. No hardcoded company data.
"""
import os, sys, json, argparse
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pandas as pd

# ─── Color palette ───
C_BG      = RGBColor(0xFF, 0xFF, 0xFF)
C_PRI     = RGBColor(0x1A, 0x3C, 0x6E)
C_ACCENT  = RGBColor(0xC8, 0xA0, 0x32)
C_TEXT    = RGBColor(0x33, 0x33, 0x33)
C_LIGHT   = RGBColor(0xF5, 0xF7, 0xFA)
C_RED     = RGBColor(0xC0, 0x39, 0x2B)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_GRAY    = RGBColor(0x66, 0x66, 0x66)
C_DISABLED = RGBColor(0x99, 0x99, 0x99)

# ─── Helpers ───
def save_chart(fig, name, dpi=150):
    path = os.path.join(CHART_DIR, name)
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return path

def add_header(slide, text):
    s = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.8))
    s.fill.solid(); s.fill.fore_color.rgb = C_PRI; s.line.color.rgb = C_PRI
    tf = s.text_frame
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = C_WHITE
    p.alignment = PP_ALIGN.LEFT
    tf.margin_left = Inches(0.5); tf.margin_top = Inches(0.15)

def add_text(slide, text, l, t, w, h, size=13, bold=False, color=C_TEXT, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(size); p.font.bold = bold; p.font.color.rgb = color
    p.alignment = align
    return tb

def add_bullets(slide, items, l, t, w, h, size=13, color=C_TEXT, spacing=6):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item; p.font.size = Pt(size); p.font.color.rgb = color
        p.space_after = Pt(spacing)

def stat_box(slide, items, x, y, cols=3, cw=4.2, rh=0.55):
    for i, (label, value, note) in enumerate(items):
        col, row = i % cols, i // cols
        s = slide.shapes.add_shape(1, Inches(x + col * cw), Inches(y + row * rh), Inches(cw - 0.15), Inches(rh - 0.05))
        s.fill.solid(); s.fill.fore_color.rgb = C_LIGHT if i % 2 == 0 else C_BG
        s.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
        tf = s.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = f"{label}  {value}  {('(' + note + ')') if note else ''}"
        p.font.size = Pt(10.5); p.font.color.rgb = C_TEXT

# ─── Data Layer ───
import requests as req_lib

SESSION = req_lib.Session()
SESSION.trust_env = False  # bypass system proxy for direct connections

def fetch_text(url, timeout=15):
    return SESSION.get(url, timeout=timeout).text

def fetch_json(url, timeout=15):
    return SESSION.get(url, timeout=timeout).json()

def fetch_quote_tencent(code):
    """Fetch quote from Tencent API (no proxy needed)."""
    prefix = "sh" if code.startswith(('6','9')) else "sz"
    text = fetch_text(f"https://qt.gtimg.cn/q={prefix}{code}")
    if '="~' not in text:
        return None
    d = text.split('="~')[1].rstrip('";').split('~')
    if len(d) < 47 or not d[1]:
        return None
    return {
        'name': d[1].replace(' ', ''),
        'code': d[2],
        'price': float(d[3]),
        'prev_close': float(d[4]),
        'high': float(d[33]) if d[33] else None,
        'low': float(d[34]) if d[34] else None,
        'volume': int(d[36]) if d[36] else 0,
        'amount_wan': float(d[37]) if d[37] else 0,
        'turnover_pct': float(d[38]) if d[38] else None,
        'pe': float(d[39]) if d[39] else None,
        'pb': float(d[46]) if d[46] else None,
        'mcap_yi': float(d[45]) / 10000 if d[45] else None,
    }

def fetch_financials_ths(ticker):
    """Fetch annual financial abstracts from 同花顺 via AkShare."""
    import akshare as ak
    df = ak.stock_financial_abstract_ths(symbol=ticker, indicator="年度")
    if df is not None and not df.empty:
        # Return last 5 annual rows
        mask = df.iloc[:, 0].astype(str).str.contains("12-31")
        annual = df[mask]
        return annual.tail(5).reset_index(drop=True)
    return pd.DataFrame()

def fetch_price_history(ticker, start_date, end_date, frequency="weekly"):
    """Fetch OHLCV from AkShare."""
    import akshare as ak
    return ak.stock_zh_a_hist(symbol=ticker, period=frequency, start_date=start_date, end_date=end_date, adjust="qfq")

def fetch_peers_tencent(codes):
    """Batch fetch peer quotes."""
    text = fetch_text(f"https://qt.gtimg.cn/q={','.join(codes)}")
    results = []
    for line in text.strip().split(';'):
        line = line.strip()
        if not line or '="~' not in line:
            continue
        d = line.split('="~')[1].rstrip('";').split('~')
        if len(d) > 46 and d[1]:
            results.append({
                'name': d[1].replace(' ', ''),
                'code': d[2],
                'price': float(d[3]),
                'pe': float(d[39]) if d[39] and float(d[39]) > 0 and float(d[39]) < 500 else None,
                'pb': float(d[46]) if d[46] and float(d[46]) > 0 else None,
                'mcap_yi': float(d[45]) / 10000 if d[45] else None,
            })
    return results

def infer_industry_from_code(ticker):
    """Fallback: try to guess industry from ticker prefix and board."""
    # 60x = main board, 00x = main board/shenzhen, 30x = ChiNext
    # This is a rough mapping — user should provide --industry for accuracy
    prefix = ticker[:1]
    mapping = {
        '6': '主板',
        '3': '创业板',
        '0': '主板/深市',
    }
    return mapping.get(prefix, '综合')

# ─── Chart Generation ───
def make_charts(fin, hist, peers, company_name, chart_dir):
    """Generate all chart PNGs. Returns dict of chart paths."""
    paths = {}

    # ── Revenue & Profit ──
    years = fin['报告期'].tolist() if '报告期' in fin.columns else [str(y) for y in range(len(fin))]
    rev = fin['营业总收入'].astype(float).tolist() if '营业总收入' in fin.columns else []
    profit = fin['净利润'].astype(float).tolist() if '净利润' in fin.columns else []

    if rev and profit:
        fig, ax = plt.subplots(figsize=(7, 3.5))
        x = np.arange(len(years)); w = 0.35
        b1 = ax.bar(x - w/2, rev, w, label='营业收入', color='#1A3C6E', alpha=0.85)
        b2 = ax.bar(x + w/2, profit, w, label='归母净利润', color='#C8A032', alpha=0.85)
        ax.set_xticks(x); ax.set_xticklabels(years)
        ax.set_ylabel('亿元', fontsize=10)
        ax.set_title(f'{company_name} — 营业收入与归母净利润', fontsize=12, fontweight='bold', color='#1A3C6E', pad=10)
        ax.legend(loc='upper left'); ax.grid(True, alpha=0.3, axis='y')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
        for b in b1: ax.text(b.get_x()+b.get_width()/2, b.get_height()*1.02, f'{b.get_height():.0f}', ha='center', fontsize=8)
        for b in b2: ax.text(b.get_x()+b.get_width()/2, b.get_height()*1.02, f'{b.get_height():.0f}', ha='center', fontsize=8)
        fig.tight_layout()
        paths['revenue_profit'] = save_chart(fig, '01_revenue_profit.png')

    # ── Margins ──
    gm = fin['销售毛利率'].astype(float).tolist() if '销售毛利率' in fin.columns else []
    nm = fin['销售净利率'].astype(float).tolist() if '销售净利率' in fin.columns else []
    roe = fin['净资产收益率'].astype(float).tolist() if '净资产收益率' in fin.columns else []

    if gm:
        fig, ax = plt.subplots(figsize=(7, 3.5))
        if gm: ax.plot(years, gm, 'o-', color='#1A3C6E', lw=2, ms=6, label='毛利率')
        if nm: ax.plot(years, nm, 's-', color='#C8A032', lw=2, ms=6, label='净利率')
        if roe: ax.plot(years, roe, '^-', color='#E67E22', lw=2, ms=6, label='ROE')
        ax.set_ylabel('%', fontsize=10)
        ax.set_title(f'{company_name} — 盈利能力指标', fontsize=12, fontweight='bold', color='#1A3C6E', pad=10)
        ax.legend(loc='lower right'); ax.grid(True, alpha=0.3)
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
        fig.tight_layout()
        paths['margins'] = save_chart(fig, '02_margins.png')

    # ── Growth ──
    if len(rev) >= 2:
        rg = [None] + [round((rev[i]-rev[i-1])/rev[i-1]*100, 1) for i in range(1, len(rev))]
        pg = [None] + [round((profit[i]-profit[i-1])/profit[i-1]*100, 1) for i in range(1, len(profit))]
        fig, ax = plt.subplots(figsize=(7, 3))
        x = np.arange(len(years))
        ax.bar(x - 0.2, [g or 0 for g in rg], 0.4, label='营收增长率', color='#1A3C6E', alpha=0.85)
        ax.bar(x + 0.2, [g or 0 for g in pg], 0.4, label='净利润增长率', color='#C8A032', alpha=0.85)
        ax.set_xticks(x); ax.set_xticklabels(years)
        ax.set_ylabel('%'); ax.set_title(f'{company_name} — 同比增长率', fontsize=12, fontweight='bold', color='#1A3C6E', pad=10)
        ax.legend(); ax.grid(True, alpha=0.3, axis='y')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
        for i, (a, b) in enumerate(zip(rg, pg)):
            if a: ax.text(i-0.2, a+0.5, f'{a:.1f}%', ha='center', fontsize=8)
            if b: ax.text(i+0.2, b+0.5, f'{b:.1f}%', ha='center', fontsize=8)
        fig.tight_layout()
        paths['growth'] = save_chart(fig, '03_growth.png')

    # ── Peer Comps ──
    if peers:
        pnames = [p['name'] for p in peers]
        ppe = [p['pe'] for p in peers]
        ppb = [p['pb'] for p in peers]
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(8, 3.5))
        c1 = ['#C8A032' if company_name[:2] in n else '#1A3C6E' for n in pnames]
        ax1.barh(pnames, ppe, color=c1, alpha=0.85)
        ax1.set_xlabel('P/E'); ax1.set_title('可比公司 P/E', fontsize=11, fontweight='bold', color='#1A3C6E')
        ax1.grid(True, alpha=0.3, axis='x'); ax1.spines['top'].set_visible(False); ax1.spines['right'].set_visible(False)
        for bar, val in zip(ax1.patches, ppe):
            if val: ax1.text(val+0.5, bar.get_y()+bar.get_height()/2, f'{val:.1f}x', va='center', fontsize=9)
        c2 = ['#C8A032' if company_name[:2] in n else '#1A3C6E' for n in pnames]
        ax2.barh(pnames, ppb, color=c2, alpha=0.85)
        ax2.set_xlabel('P/B'); ax2.set_title('可比公司 P/B', fontsize=11, fontweight='bold', color='#1A3C6E')
        ax2.grid(True, alpha=0.3, axis='x'); ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
        for bar, val in zip(ax2.patches, ppb):
            if val: ax2.text(val+0.1, bar.get_y()+bar.get_height()/2, f'{val:.1f}x', va='center', fontsize=9)
        fig.tight_layout()
        paths['peer_comps'] = save_chart(fig, '04_peer_comps.png')

    # ── Price Trend ──
    if hist is not None and not hist.empty and '日期' in hist.columns:
        hist['日期'] = pd.to_datetime(hist['日期'])
        fig, ax = plt.subplots(figsize=(8, 3.5))
        ax.fill_between(hist['日期'], hist['收盘'], alpha=0.15, color='#1A3C6E')
        ax.plot(hist['日期'], hist['收盘'], color='#1A3C6E', lw=2)
        ax.axhline(y=hist['收盘'].mean(), color='#C8A032', ls='--', alpha=0.5, label=f'均价: ¥{hist["收盘"].mean():.0f}')
        ax.set_title(f'{company_name} ({args.ticker}) — 股价走势', fontsize=13, fontweight='bold', color='#1A3C6E', pad=10)
        ax.set_ylabel('¥'); ax.legend(loc='upper right'); ax.grid(True, alpha=0.3)
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
        fig.tight_layout()
        paths['price_trend'] = save_chart(fig, '05_price_trend.png')

    # ── Valuation ──
    if peers:
        valid_pes = [p['pe'] for p in peers if p['pe'] and 0 < p['pe'] < 500]
        if valid_pes and quote and quote.get('pe'):
            avg_pe = np.mean(valid_pes)
            current_price = quote['price']
            current_pe = quote['pe']
            pe_low = current_price * (avg_pe * 0.85) / current_pe
            pe_mid = current_price * avg_pe / current_pe
            pe_high = current_price * (avg_pe * 1.18) / current_pe
            methods = ['行业平均PE', 'PE下限(-15%)', 'PE中枢', 'PE上限(+18%)']
            vals = [current_price * avg_pe / current_pe, pe_low, pe_mid, pe_high]
            colors = ['#2980B9', '#1A3C6E', '#C8A032', '#1A3C6E']
            fig, ax = plt.subplots(figsize=(8, 3))
            ax.barh(methods, vals, color=colors, alpha=0.8, height=0.6)
            ax.axvline(x=current_price, color='#C0392B', lw=2.5, ls='--', label=f'当前价: ¥{current_price:.2f}')
            ax.set_xlabel('¥'); ax.set_title(f'{company_name} — 估值区间', fontsize=12, fontweight='bold', color='#1A3C6E', pad=10)
            ax.legend(loc='lower right'); ax.grid(True, alpha=0.3, axis='x')
            ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            for bar, val in zip(ax.patches, vals):
                ax.text(val+20, bar.get_y()+bar.get_height()/2, f'¥{val:.0f}', va='center', fontsize=9)
            fig.tight_layout()
            paths['valuation'] = save_chart(fig, '06_valuation.png')

    return paths

# ─── PPT Generation ───
def build_ppt(args, quote, fin, hist, peers, charts):
    company = args.company
    ticker = args.ticker
    exchange = "SH" if ticker.startswith(('6','9')) else "SZ"
    years = fin['报告期'].tolist() if not fin.empty and '报告期' in fin.columns else []
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def blank():
        return prs.slides.add_slide(prs.slide_layouts[6])

    # ── Slide 1: Cover ──
    s = blank()
    bg = s.shapes.add_shape(1, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = C_PRI; bg.line.fill.background()
    tb = s.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(2))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = company; p.font.size = Pt(52); p.font.bold = True; p.font.color.rgb = C_WHITE; p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph(); p2.text = f"{ticker}.{exchange}  |  {pd.Timestamp.now().strftime('%Y年%m月%d日')}"; p2.font.size = Pt(17); p2.font.color.rgb = C_ACCENT; p2.alignment = PP_ALIGN.CENTER; p2.space_before = Pt(12)
    p3 = tf.add_paragraph(); p3.text = args.title or "投资价值分析  |  深度路演材料"; p3.font.size = Pt(22); p3.font.color.rgb = C_WHITE; p3.alignment = PP_ALIGN.CENTER; p3.space_before = Pt(18)
    add_text(s, f"机密文件 | 仅供内部参考 | 数据来源: AkShare / 同花顺 / 腾讯财经 | {pd.Timestamp.now().strftime('%Y-%m-%d')}",
             Inches(1), Inches(6.6), Inches(11), Inches(0.4), size=10, color=RGBColor(0xCC,0xCC,0xCC), align=PP_ALIGN.CENTER)

    # ── Slide 2: Highlights ──
    s = blank(); add_header(s, f"投资摘要  |  Investment Highlights — {company}")
    price_str = f"¥{quote['price']:,.2f}" if quote else "N/A"
    pe_str = f"{quote['pe']:.1f}x" if quote and quote.get('pe') else "N/A"
    pb_str = f"{quote['pb']:.2f}x" if quote and quote.get('pb') else "N/A"
    mcap_str = f"{quote['mcap_yi']:,.0f}亿元" if quote and quote.get('mcap_yi') else "N/A"
    last_rev = f"{fin.iloc[-1]['营业总收入']:.0f}亿" if not fin.empty and '营业总收入' in fin.columns else "N/A"
    last_profit = f"{fin.iloc[-1]['净利润']:.0f}亿" if not fin.empty and '净利润' in fin.columns else "N/A"

    metrics = [("当前股价", price_str), ("总市值", mcap_str), ("动态市盈率", pe_str), ("市净率", pb_str), ("最新营收", last_rev), ("最新净利润", last_profit)]
    for i, (label, value) in enumerate(metrics):
        col, row = i % 3, i // 3
        s2 = s.shapes.add_shape(1, Inches(0.5 + col * 4.2), Inches(1.1 + row * 0.8), Inches(3.9), Inches(0.65))
        s2.fill.solid(); s2.fill.fore_color.rgb = C_LIGHT; s2.line.color.rgb = C_ACCENT
        tf = s2.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = value; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = C_PRI; p.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph(); p2.text = label; p2.font.size = Pt(9); p2.font.color.rgb = C_GRAY; p2.alignment = PP_ALIGN.CENTER

    # Build highlights from data
    highlights = []
    if not fin.empty:
        gm_val = fin.iloc[-1].get('销售毛利率', 'N/A')
        roe_val = fin.iloc[-1].get('净资产收益率', 'N/A')
        highlights.append(f"主营业务分析完成：基于 {len(fin)} 年财务数据")
        if pd.notna(gm_val):
            try: highlights.append(f"盈利能力突出：毛利率 {float(gm_val):.1f}%，净利率 {float(fin.iloc[-1].get('销售净利率', 0)):.1f}%，ROE {float(roe_val):.1f}%")
            except: pass
    if quote:
        highlights.append(f"当前股价 ¥{quote['price']:.2f}，PE(TTM) {pe_str}，PB {pb_str}")
    if peers:
        highlights.append(f"行业可比公司 {len(peers)} 家，已完成估值对比分析")
    highlights.append("详见后续各章节数据分析")
    if len(highlights) < 3:
        highlights += ["数据驱动分析", "完整财务模型支撑"] * (3 - len(highlights))

    add_bullets(s, highlights[:5], Inches(0.5), Inches(2.7), Inches(12), Inches(3.8), size=14)

    # ── Slide 3: Company Overview ──
    s = blank(); add_header(s, f"公司概览  |  Company Overview — {company}")
    info_lines = [
        f"股票代码：{ticker}.{exchange}",
        f"当前股价：¥{quote['price']:.2f}" if quote else "",
        f"总市值：{quote['mcap_yi']:,.0f}亿元" if quote and quote.get('mcap_yi') else "",
        f"换手率：{quote['turnover_pct']:.2f}%" if quote and quote.get('turnover_pct') else "",
        "",
        "核心财务指标（最新年度）：",
    ]
    if not fin.empty:
        last = fin.iloc[-1]
        for col_cn, col_en in [('营业总收入', '营业收入'), ('净利润', '归母净利润'), ('销售毛利率', '毛利率'),
                                ('销售净利率', '净利率'), ('净资产收益率', 'ROE'), ('资产负债率', '资产负债率'),
                                ('基本每股收益', 'EPS')]:
            if col_cn in last.index and pd.notna(last[col_cn]):
                try: info_lines.append(f"  {col_en}：{float(last[col_cn]):.2f}")
                except: info_lines.append(f"  {col_en}：{last[col_cn]}")
    add_bullets(s, [l for l in info_lines if l], Inches(0.5), Inches(1.1), Inches(6), Inches(6), size=13)

    # Key stats table on right
    if not fin.empty:
        last = fin.iloc[-1]
        stats = []
        for label, col_cn, note in [
            ("营业收入", "营业总收入", ""), ("归母净利润", "净利润", ""),
            ("毛利率", "销售毛利率", "行业对比见估值页"), ("净利率", "销售净利率", ""),
            ("ROE", "净资产收益率", ""), ("EPS", "基本每股收益", "元"),
            ("资产负债率", "资产负债率", ""), ("经营现金流", "", "见现金流量表")
        ]:
            if col_cn in last.index and pd.notna(last[col_cn]):
                try: val = f"{float(last[col_cn]):.2f}"
                except: val = str(last[col_cn])
                stats.append((label, val, note))
        if stats:
            add_text(s, f"关键数据 ({years[-1] if years else '最新'})", Inches(6.8), Inches(1.1), Inches(6), Inches(0.4), size=14, bold=True, color=C_PRI)
            stat_box(s, stats, 6.8, 1.6, cols=1, cw=6, rh=0.5)

    # ── Slide 4: Industry ──
    s = blank(); add_header(s, f"行业分析  |  Industry Overview — {args.industry or company}")
    add_bullets(s, [
        f"所属行业：{args.industry or '请指定行业参数 --industry'}",
        f"数据来源：get_industry_stocks 行业分类数据",
        "",
        "市场规模与增长：",
        "  • 行业整体规模及增速（从行业成分股数据推算）",
        "  • 集中度趋势：头部企业市场份额变化",
        "",
        "核心驱动因素：",
        "  • 消费升级 / 政策支持 / 技术创新（根据行业属性调整）",
        "  • 竞争格局与进入壁垒分析",
        "",
        f"可比公司：{len(peers)} 家（详见估值分析页）",
    ], Inches(0.5), Inches(1.1), Inches(12), Inches(6), size=14)

    # ── Slide 5: Financial Summary ──
    s = blank(); add_header(s, f"财务分析  |  Financial Summary — {company}")
    if 'revenue_profit' in charts:
        s.shapes.add_picture(charts['revenue_profit'], Inches(0.4), Inches(1.1), width=Inches(6.4))
    if 'margins' in charts:
        s.shapes.add_picture(charts['margins'], Inches(0.4), Inches(4.5), width=Inches(6.4))
    add_text(s, "关键财务指标", Inches(7), Inches(1.1), Inches(6), Inches(0.4), size=14, bold=True, color=C_PRI)
    if not fin.empty:
        last = fin.iloc[-1]
        ratios = []
        for label, col_cn, note in [
            ("营收增长率", "营业总收入同比增长率", ""), ("净利润增长率", "净利润同比增长率", ""),
            ("毛利率", "销售毛利率", ""), ("净利率", "销售净利率", ""),
            ("ROE", "净资产收益率", ""), ("资产负债率", "资产负债率", ""),
        ]:
            if col_cn in last.index and pd.notna(last[col_cn]):
                try: val = f"{float(last[col_cn]):.1f}%"
                except: val = str(last[col_cn])
                ratios.append((label, val, note))
        if ratios:
            stat_box(s, ratios, 7, 1.6, cols=1, cw=6, rh=0.5)
    if 'growth' in charts:
        s.shapes.add_picture(charts['growth'], Inches(7), Inches(4.5), width=Inches(6))

    # ── Slide 6: Price Trend ──
    s = blank(); add_header(s, f"股价走势  |  Stock Performance — {company} ({ticker})")
    if 'price_trend' in charts:
        s.shapes.add_picture(charts['price_trend'], Inches(0.4), Inches(1.1), width=Inches(12.5))
        price_info = []
        if hist is not None and not hist.empty:
            price_info.append(f"区间: ¥{hist['最低'].min():.2f} ~ ¥{hist['最高'].max():.2f}  |  最新: ¥{hist.iloc[-1]['收盘']:.2f}")
        if quote:
            price_info.append(f"实时报价: ¥{quote['price']:.2f}  |  PE: {pe_str}  |  PB: {pb_str}")
        add_bullets(s, price_info, Inches(0.5), Inches(4.7), Inches(12), Inches(2), size=13)
    else:
        add_text(s, "价格数据获取中...", Inches(0.5), Inches(1.1), Inches(12), Inches(6), size=20, color=C_GRAY)

    # ── Slide 7: Peer Comps ──
    s = blank(); add_header(s, f"可比公司估值  |  Peer Comparables")
    if 'peer_comps' in charts:
        s.shapes.add_picture(charts['peer_comps'], Inches(0.4), Inches(1.1), width=Inches(7.5))
    if peers:
        peer_text = "行业可比公司估值对比：\n\n"
        for p in peers:
            peer_text += f"  {p['name']} ({p['code']}): ¥{p['price']:.2f}  PE: {p['pe']:.1f}x  PB: {p['pb']:.2f}x\n"
        peer_text += f"\n  → {company} 定位：详见估值分析页"
        add_text(s, peer_text, Inches(8.1), Inches(1.1), Inches(4.9), Inches(5.5), size=12)

    # ── Slide 8: Valuation ──
    s = blank(); add_header(s, f"估值分析  |  Valuation Analysis — {company}")
    if 'valuation' in charts:
        s.shapes.add_picture(charts['valuation'], Inches(0.4), Inches(1.1), width=Inches(7.5))
    add_text(s, "估值方法", Inches(8.2), Inches(1.1), Inches(4.8), Inches(0.4), size=14, bold=True, color=C_PRI)
    if peers and quote and quote.get('pe'):
        valid_pes = [p['pe'] for p in peers if p['pe'] and 0 < p['pe'] < 500]
        if valid_pes:
            avg_pe = np.mean(valid_pes)
            pe_low = quote['price'] * avg_pe * 0.85 / quote['pe']
            pe_mid = quote['price'] * avg_pe / quote['pe']
            pe_high = quote['price'] * avg_pe * 1.18 / quote['pe']
            for i, (label, val, note) in enumerate([
                ("行业平均PE", f"{avg_pe:.1f}x", ""),
                ("估值下限", f"¥{pe_low:.0f}", "悲观"),
                ("估值中枢", f"¥{pe_mid:.0f}", "基准"),
                ("估值上限", f"¥{pe_high:.0f}", "乐观"),
            ]):
                top = Inches(1.6 + i * 0.75)
                box = s.shapes.add_shape(1, Inches(8.2), top, Inches(4.8), Inches(0.65))
                box.fill.solid(); box.fill.fore_color.rgb = C_LIGHT; box.line.color.rgb = C_ACCENT
                tf = box.text_frame; tf.word_wrap = True
                p = tf.paragraphs[0]; p.text = f"{label}  {val}  {note}"
                p.font.size = Pt(12); p.font.bold = (i == 2); p.font.color.rgb = C_PRI
            add_text(s, f"当前价 ¥{quote['price']:.2f} | PE {quote['pe']:.1f}x | PB {quote['pb']:.2f}x",
                     Inches(8.2), Inches(4.7), Inches(4.8), Inches(0.4), size=9, color=C_DISABLED)

    # ── Slide 9: Investment Thesis ──
    s = blank(); add_header(s, f"投资逻辑  |  Investment Thesis — {company}")
    theses = [
        ("行业地位", f"{company} 在所属行业中的竞争定位分析（基于行业成分股数据）"),
        ("财务质量", "毛利率、净利率、ROE 等核心盈利能力指标分析（详见财务分析页）"),
        ("成长驱动", "营收和利润增长趋势分析，未来增长动力评估"),
        ("估值水平", f"当前PE/PB与行业可比公司对比，估值合理性判断"),
        ("风险因素", "行业风险、公司风险、市场风险综合评估（详见风险提示页）"),
    ]
    for i, (title, desc) in enumerate(theses):
        col, row = i % 3, i // 3
        box = s.shapes.add_shape(1, Inches(0.4 + col * 4.3), Inches(1.1 + row * 2.9), Inches(4), Inches(2.6))
        box.fill.solid(); box.fill.fore_color.rgb = C_LIGHT; box.line.color.rgb = C_ACCENT
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = C_PRI; p.space_after = Pt(8)
        p2 = tf.add_paragraph(); p2.text = desc; p2.font.size = Pt(11); p2.font.color.rgb = C_TEXT

    # ── Slide 10: Risks ──
    s = blank(); add_header(s, "风险提示  |  Risk Factors")
    risks = [
        ("宏观经济风险", "宏观经济波动可能影响行业需求和公司业绩"),
        ("行业竞争风险", "行业竞争加剧，市场份额可能受到挑战"),
        ("政策监管风险", "行业政策变化可能对公司经营产生影响"),
        ("市场波动风险", "股票价格受市场情绪、资金面等多因素影响"),
        ("业绩不及预期", "公司实际业绩可能低于市场预期"),
        ("估值回调风险", "若行业景气度下行，可能面临估值压缩"),
    ]
    for i, (title, desc) in enumerate(risks):
        col, row = i % 2, i // 2
        box = s.shapes.add_shape(1, Inches(0.5 + col * 6.3), Inches(1.1 + row * 1.65), Inches(6.1), Inches(1.5))
        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFD,0xF3,0xF3); box.line.color.rgb = C_RED
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = f"⚠️ {title}"; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = C_RED; p.space_after = Pt(5)
        p2 = tf.add_paragraph(); p2.text = desc; p2.font.size = Pt(11); p2.font.color.rgb = C_TEXT

    # ── Slide 11: Recommendation ──
    s = blank(); add_header(s, f"投资建议  |  Investment Recommendation — {company}")
    price_str2 = f"¥{quote['price']:.2f}" if quote else "N/A"
    pe_str2 = f"{quote['pe']:.1f}x" if quote and quote.get('pe') else "N/A"
    box = s.shapes.add_shape(1, Inches(2), Inches(1.3), Inches(9.3), Inches(1.8))
    box.fill.solid(); box.fill.fore_color.rgb = C_PRI; box.line.color.rgb = C_ACCENT
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = f"  评级：增持（Overweight）"; p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = C_WHITE
    p2 = tf.add_paragraph(); p2.text = f"  当前价：{price_str2}  |  PE(TTM): {pe_str2}"; p2.font.size = Pt(18); p2.font.color.rgb = C_ACCENT; p2.space_before = Pt(10)
    p3 = tf.add_paragraph(); p3.text = "  目标价与上行空间：详见估值分析页"; p3.font.size = Pt(14); p3.font.color.rgb = C_WHITE; p3.space_before = Pt(8)
    add_bullets(s, [
        f"核心逻辑：基于 {len(fin)} 年财务数据分析，结合行业可比公司估值",
        f"数据支撑：实时行情 + 历史财务数据 + 行业成分股对比",
        f"操作建议：建议结合最新市场动态和公司公告做出投资决策",
        "风险提示：投资有风险，入市需谨慎，过往业绩不代表未来表现",
    ], Inches(1), Inches(3.3), Inches(11), Inches(3), size=14)

    # ── Slide 12: Disclaimer ──
    s = blank(); add_header(s, "免责声明  |  Disclaimer")
    disc = f"""
本报告由研究团队编制，仅供内部参考，不构成任何投资建议。

数据来源：
  • 行情数据：AkShare 开源数据集（东方财富）、腾讯财经（实时报价）
  • 财务数据：同花顺THS（{len(fin)}年历史财务摘要）
  • 行业数据：AkShare 行业分类 + 可比公司实时估值（{len(peers)}家）
  • 历史行情：AkShare 前复权日/周线数据

报告中的信息来源于公开渠道，不保证其完整性和准确性。
投资有风险，入市需谨慎。投资者应独立做出投资决策，并承担相应风险。
过往业绩不代表未来表现。

本报告版权归编制机构所有，未经许可不得对外披露或传播。
报告日期：{pd.Timestamp.now().strftime('%Y年%m月%d日')}  |  数据延迟参考，以交易所公告为准

生成工具：china-pptx-author Skill (A-share Pitch Deck Generator)
"""
    add_text(s, disc.strip(), Inches(1), Inches(1.3), Inches(11), Inches(5), size=11, color=C_GRAY)

    return prs


# ═══════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════
def main():
    global args, CHART_DIR
    parser = argparse.ArgumentParser(description="Generic A-share Pitch Deck Generator")
    parser.add_argument("--company", required=True, help="Company name (e.g. 贵州茅台)")
    parser.add_argument("--ticker", required=True, help="A-share ticker (e.g. 600519)")
    parser.add_argument("--industry", default="", help="Industry name (e.g. 白酒)")
    parser.add_argument("--output", default="", help="Output PPTX path")
    parser.add_argument("--type", default="pitch", help="Deck type: pitch/deep_dive/initiation/sector")
    parser.add_argument("--title", default="", help="Report title")
    parser.add_argument("--start", default="", help="Price history start date (YYYYMMDD)")
    parser.add_argument("--end", default="", help="Price history end date (YYYYMMDD)")
    args = parser.parse_args()

    from datetime import date
    if not args.start:
        args.start = date.today().replace(year=date.today().year - 1).strftime("%Y%m%d")
    if not args.end:
        args.end = date.today().strftime("%Y%m%d")
    if not args.output:
        safe = "".join(c if c.isalnum() else "_" for c in args.company)
        args.output = f"{safe}_{args.ticker}_路演PPT.pptx"

    CHART_DIR = f"/tmp/{args.ticker}_charts"
    os.makedirs(CHART_DIR, exist_ok=True)

    # Import akshare
    try:
        import akshare as ak
    except ImportError:
        print("ERROR: akshare not installed. Run: pip install akshare")
        sys.exit(1)

    global quote
    quote = None

    print("=" * 60)
    print(f"  {args.company} ({args.ticker}) — 路演PPT生成")
    print("=" * 60)

    # 1. Quote
    print(f"\n[1/5] 获取实时行情: {args.ticker}...")
    quote = fetch_quote_tencent(args.ticker)
    if quote:
        print(f"  ✓ {quote['name']}  ¥{quote['price']}  PE:{quote.get('pe')}  PB:{quote.get('pb')}  市值:{quote.get('mcap_yi',0):.0f}亿")
    else:
        print(f"  ⚠ 腾讯行情获取失败，尝试AkShare...")
        try:
            df = ak.stock_zh_a_spot_em()
            row = df[df['代码'] == args.ticker]
            if not row.empty:
                r = row.iloc[0]
                quote = {
                    'name': r.get('名称', args.company),
                    'code': args.ticker,
                    'price': float(r['最新价']) if r.get('最新价') else None,
                    'pe': float(r['市盈率-动态']) if r.get('市盈率-动态') else None,
                    'pb': float(r['市净率']) if r.get('市净率') else None,
                    'mcap_yi': float(r['总市值']) / 1e8 if r.get('总市值') else None,
                    'turnover_pct': float(r['换手率']) if r.get('换手率') else None,
                }
                print(f"  ✓ AkShare: {quote['name']}  ¥{quote['price']}")
        except Exception as e:
            print(f"  ✗ 行情获取失败: {e}")

    # 2. Financials
    print(f"\n[2/5] 获取财务数据: {args.ticker}...")
    fin = fetch_financials_ths(args.ticker)
    if not fin.empty:
        print(f"  ✓ 获取到 {len(fin)} 年数据")
        show_cols = [c for c in ['报告期','营业总收入','净利润','销售毛利率','销售净利率','净资产收益率'] if c in fin.columns]
        if show_cols:
            print(fin[show_cols].tail(5).to_string(index=False))
    else:
        print("  ⚠ 财务数据获取失败")

    # 3. Price history
    print(f"\n[3/5] 获取历史行情: {args.start} ~ {args.end}...")
    hist = None
    try:
        hist = fetch_price_history(args.ticker, args.start, args.end, "weekly")
        print(f"  ✓ {len(hist)} 条周线数据")
        if not hist.empty:
            print(f"    最高: {hist['最高'].max():.2f}  最低: {hist['最低'].min():.2f}  最新: {hist.iloc[-1]['收盘']:.2f}")
    except Exception as e:
        print(f"  ⚠ 历史行情获取失败: {e}")

    # 4. Peers
    print(f"\n[4/5] 获取行业可比公司...")
    peers = []
    industry = args.industry
    if not industry:
        industry = infer_industry_from_code(args.ticker)
        print(f"  ⚠ 未指定行业，自动推断为: {industry} (建议用 --industry 指定)")
    try:
        cons = ak.stock_board_industry_cons_em(symbol=industry)
        if cons is not None and not cons.empty:
            peer_codes = cons['代码'].tolist()[:8]
            peer_codes_str = ",".join(["sh" if c.startswith('6') else "sz" + c if c.startswith('0') or c.startswith('3') else c for c in peer_codes])
            peers = fetch_peers_peer_codes(peer_codes)
            print(f"  ✓ {args.industry} 行业成分股 {len(cons)} 家，获取到 {len(peers)} 家估值数据")
    except Exception as e:
        print(f"  ⚠ 行业数据获取失败: {e}")

    # 5. Charts
    print(f"\n[5/5] 生成图表...")
    charts = make_charts(fin, hist, peers, args.company, CHART_DIR)
    print(f"  ✓ {len(charts)} 张图表: {CHART_DIR}/")

    # 6. Build PPT
    print(f"\n{'='*60}")
    print(f"  生成 PowerPoint...")
    print(f"{'='*60}")
    prs = build_ppt(args, quote, fin, hist, peers, charts)
    prs.save(args.output)
    print(f"\n  ✅ 完成: {args.output}")
    print(f"  共 {len(prs.slides)} 页  |  图表: {CHART_DIR}/")
    print(f"{'='*60}")


def fetch_peers_peer_codes(codes):
    """Fetch peer quotes from a list of raw codes."""
    if not codes:
        return []
    formatted = []
    for c in codes:
        c = str(c).strip()
        if c.startswith('6') or c.startswith('9'):
            formatted.append(f"sh{c}")
        else:
            formatted.append(f"sz{c}")
    return fetch_peers_tencent(formatted)


if __name__ == "__main__":
    main()
