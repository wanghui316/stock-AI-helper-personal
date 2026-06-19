# -*- coding: utf-8 -*-
"""
A股持仓/备选股每日量化评分系统（混合数据源版）
数据源组合：
  1. 实时行情（价格、成交量）：雪球 pysnowball
  2. 历史K线（用于趋势和量价分析）：baostock
  3. TTM PE：自动获取失败后通过人机交互输入

使用方法：
  1. 安装依赖：pip install pysnowball baostock
  2. 直接运行，按提示输入PE值
  3. 输出Excel评分报告
"""

import pysnowball as ball
import baostock as bs
import pandas as pd
import numpy as np
import warnings
import os
import sys
from datetime import datetime, timedelta

warnings.filterwarnings('ignore')

# ==================== 自动路径识别 ====================
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

print(f"📁 脚本所在目录：{script_dir}")

# ==================== Token配置 ====================
TOKEN = "xq_a_token=3a22b8d793271ca17cd0a1eb8d1f28a3a8d96fb3; u=7772226752"
ball.set_token(TOKEN)

# 测试Token
try:
    test = ball.quotec('SZ000001')
    if test and test.get('data'):
        print("✅ 雪球Token配置成功！")
    else:
        print("⚠️ Token验证失败，请检查")
except Exception as e:
    print(f"❌ Token配置失败：{e}")
    sys.exit(1)

# ==================== 用户配置区 ====================

STOCK_POOL = [
    {'code': 'sh.600522', 'name': '中天科技', 'type': 'Ⅱ', 'level': '🔵', 'sector': '通信设备', 'status': '持仓'},
    {'code': 'sh.601899', 'name': '紫金矿业', 'type': 'Ⅰ', 'level': '🟢', 'sector': '有色金属', 'status': '持仓'},
    {'code': 'sh.600089', 'name': '特变电工', 'type': 'Ⅰ/Ⅱ', 'level': '🔵', 'sector': '电力设备', 'status': '持仓'},
    {'code': 'sh.600143', 'name': '金发科技', 'type': 'Ⅰ', 'level': '🟡', 'sector': '化工', 'status': '持仓'},
    {'code': 'sz.000338', 'name': '潍柴动力', 'type': 'Ⅰ', 'level': '🔵', 'sector': '汽车', 'status': '持仓'},
    {'code': 'sz.002340', 'name': '格林美', 'type': 'Ⅱ', 'level': '🟢', 'sector': '环保', 'status': '持仓'},
    {'code': 'sz.002472', 'name': '双环传动', 'type': 'Ⅱ', 'level': '🔵', 'sector': '机械', 'status': '持仓'},
    {'code': 'sz.002698', 'name': '博实股份', 'type': 'Ⅱ', 'level': '🟢', 'sector': '机械', 'status': '持仓'},
    {'code': 'sz.300007', 'name': '汉威科技', 'type': 'Ⅲ', 'level': '🟢', 'sector': '仪器仪表', 'status': '持仓'},
    {'code': 'sz.300014', 'name': '亿纬锂能', 'type': 'Ⅱ', 'level': '🔵', 'sector': '电力设备', 'status': '备选'},
    {'code': 'sz.002850', 'name': '科达利', 'type': 'Ⅱ', 'level': '🔵', 'sector': '电力设备', 'status': '备选'},
    {'code': 'sz.300073', 'name': '当升科技', 'type': 'Ⅱ', 'level': '🔵', 'sector': '电力设备', 'status': '备选'},
    {'code': 'sh.605117', 'name': '德业股份', 'type': 'Ⅱ', 'level': '🔵', 'sector': '电力设备', 'status': '备选'},
    {'code': 'sz.002709', 'name': '天赐材料', 'type': 'Ⅱ', 'level': '🟢', 'sector': '化工', 'status': '备选'},
]

SECTOR_PE = {
    '仪器仪表': 45.0, '通信设备': 55.0, '有色金属': 21.44,
    '电力设备': 36.66, '化工': 25.0, '汽车': 25.0, '环保': 30.0, '机械': 30.0,
}

POLICY_BASE_SCORE = {'🔵': 85, '🟢': 65, '🟡': 45}

PE_RANGE = {
    'Ⅰ': {'min': 0, 'max': 20},
    'Ⅱ': {'min': 25, 'max': 40},
    'Ⅲ': {'min': 0, 'max': 60},
    'Ⅰ/Ⅱ': {'min': 20, 'max': 30},
}

WEIGHTS = {
    'valuation': 0.40,
    'policy': 0.15,
    'rpe': 0.10,
    'trend': 0.20,
    'momentum': 0.15,
}

OUTPUT_PREFIX = 'stock_analysis'


# ==================== 数据获取函数 ====================

def get_hist_data_baostock(code, days=30):
    """使用baostock获取历史K线（近30天）"""
    try:
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days + 10)).strftime('%Y-%m-%d')
        rs = bs.query_history_k_data_plus(
            code,
            "date,close,volume",
            start_date=start_date,
            end_date=end_date,
            frequency="d",
            adjustflag="2"
        )
        if rs.error_code != '0':
            return None
        data_list = []
        while (rs.error_code == '0') & rs.next():
            data_list.append(rs.get_row_data())
        if not data_list:
            return None
        df = pd.DataFrame(data_list, columns=rs.fields)
        df['日期'] = pd.to_datetime(df['date'])
        df['收盘'] = df['close'].astype(float)
        df['成交量'] = df['volume'].astype(float)
        df = df.sort_values('日期')
        return df
    except Exception as e:
        print(f"  ⚠️ baostock K线获取失败: {e}")
    return None


def get_quote_data_xueqiu(code):
    """使用雪球获取实时行情数据"""
    try:
        # 转换代码格式：sh.600522 -> SH600522
        code_xq = code.replace('.', '').upper()
        data = ball.quotec(code_xq)
        if data and data.get('data') and len(data['data']) > 0:
            quote = data['data'][0]
            pe_ttm = quote.get('pe_ttm') or quote.get('pe') or quote.get('pe_roe') or 0
            return {
                'price': quote.get('current', 0),
                'change_pct': quote.get('percent', 0),
                'volume': quote.get('volume', 0),
                'pe_ttm': pe_ttm if pe_ttm > 0 else 0,
                'name': quote.get('name', ''),
            }
    except Exception as e:
        print(f"  ⚠️ 雪球行情获取失败: {e}")
    return None


def get_pe_value_interactive(code, stock_name, sector, pe_cache):
    """
    PE获取流程：自动获取 → 人机交互输入 → 板块均值（兜底）
    """
    # 1. 如果本次运行已输入过，直接复用缓存
    if code in pe_cache:
        print(f"  📋 复用本次输入PE: {pe_cache[code]}")
        return pe_cache[code]

    # 2. 尝试从雪球接口获取
    quote = get_quote_data_xueqiu(code)
    if quote and quote.get('pe_ttm', 0) > 0:
        pe_ttm = quote['pe_ttm']
        print(f"  📋 自动获取PE: {pe_ttm}")
        pe_cache[code] = pe_ttm
        return pe_ttm

    # 3. 自动获取失败 → 人机交互输入
    print(f"  ❓ 无法自动获取 {stock_name} 的TTM PE")
    while True:
        user_input = input(f"  📝 请输入 {code} {stock_name} 的TTM PE（回车使用板块均值）: ")
        if user_input.strip() == "":
            pe_ttm = SECTOR_PE.get(sector, 30.0)
            print(f"  📋 使用板块均值: {pe_ttm}")
            pe_cache[code] = pe_ttm
            return pe_ttm
        try:
            pe_val = float(user_input)
            if pe_val > 0:
                print(f"  📋 手动输入PE: {pe_val}")
                pe_cache[code] = pe_val
                return pe_val
            else:
                print("  ⚠️ PE必须为正数，请重新输入")
        except ValueError:
            print("  ⚠️ 请输入有效数字，或直接回车跳过")


def calc_macd(close):
    """计算MACD金叉"""
    if len(close) < 26:
        return False
    exp1 = close.ewm(span=12, adjust=False).mean()
    exp2 = close.ewm(span=26, adjust=False).mean()
    dif = exp1 - exp2
    dea = dif.ewm(span=9, adjust=False).mean()
    if len(dif) < 2:
        return False
    return dif.iloc[-1] > dea.iloc[-1] and dif.iloc[-2] <= dea.iloc[-2]


def calc_scores(code, hist, quote_data, sector, stock_info, pe_cache):
    """计算五维评分"""
    if quote_data is None or quote_data['price'] == 0:
        return None

    scores = {}
    price = quote_data['price']
    volume = quote_data['volume']
    stock_name = stock_info['name']

    # 获取PE（自动 → 人机交互 → 板块均值）
    pe_ttm = get_pe_value_interactive(code, stock_name, sector, pe_cache)

    sector_pe = SECTOR_PE.get(sector, 30.0)
    stype = stock_info['type']
    slevel = stock_info['level']

    # 估值分
    pe_range = PE_RANGE.get(stype, {'min': 0, 'max': 40})
    if pe_ttm and pe_ttm > 0:
        pe_min, pe_max = pe_range['min'], pe_range['max']
        if pe_ttm <= pe_min * 0.8:
            val_score = 95
        elif pe_ttm <= pe_min:
            val_score = 85
        elif pe_ttm <= (pe_min + pe_max) / 2:
            val_score = 70
        elif pe_ttm <= pe_max:
            val_score = 55
        elif pe_ttm <= pe_max * 1.1:
            val_score = 35
        else:
            val_score = 15
        if stype == 'Ⅲ' and pe_ttm > 60:
            val_score = min(val_score, 40)
    else:
        val_score = 50
    scores['valuation'] = val_score

    # 政策分
    scores['policy'] = POLICY_BASE_SCORE.get(slevel, 50)

    # RPE分
    if pe_ttm and pe_ttm > 0 and sector_pe > 0:
        rpe = pe_ttm / sector_pe
        if rpe < 0.70:
            rpe_score = 95
        elif rpe < 0.90:
            rpe_score = 80
        elif rpe <= 1.10:
            rpe_score = 60
        elif rpe <= 1.40:
            rpe_score = 40
        else:
            rpe_score = 20
    else:
        rpe_score = 50
    scores['rpe'] = rpe_score
    scores['rpe_value'] = pe_ttm / sector_pe if (pe_ttm and pe_ttm > 0 and sector_pe > 0) else None

    # 趋势分（使用baostock K线）
    if hist is not None and len(hist) >= 20:
        close = hist['收盘']
        ma20 = close.rolling(20).mean()
        latest_ma20 = ma20.iloc[-1]
        above_ma = price > latest_ma20
        macd_golden = calc_macd(close)
        if above_ma and macd_golden:
            trend_score = 95
        elif above_ma and not macd_golden:
            trend_score = 75
        elif not above_ma and macd_golden:
            trend_score = 55
        else:
            trend_score = 35
    else:
        trend_score = 50
    scores['trend'] = trend_score

    # 量价分（双二阶，使用baostock K线）
    if hist is not None and len(hist) >= 8:
        v = hist['成交量']
        p = hist['收盘']
        v7 = v.rolling(7).mean()
        p7 = p.rolling(7).mean()
        dv = (v7 - v7.shift(1)) / v7.shift(1) * 100
        dp = (p7 - p7.shift(1)) / p7.shift(1) * 100
        d2v = dv - dv.shift(1)
        d2p = dp - dp.shift(1)
        latest_d2v = d2v.iloc[-1] if not pd.isna(d2v.iloc[-1]) else 0
        latest_d2p = d2p.iloc[-1] if not pd.isna(d2p.iloc[-1]) else 0
        if latest_d2v > 0 and latest_d2p > 0:
            momentum_score = 95
            signal = '🟢共振做多'
        elif latest_d2v > 0 and latest_d2p < 0:
            momentum_score = 60
            signal = '🟡量热价冷'
        elif latest_d2v < 0 and latest_d2p > 0:
            momentum_score = 60
            signal = '🟡价热量冷'
        else:
            momentum_score = 30
            signal = '🔴共振减速'
        scores['momentum'] = momentum_score
        scores['signal'] = signal
        scores['d2v'] = round(latest_d2v, 2)
        scores['d2p'] = round(latest_d2p, 2)
    else:
        momentum_score = 50
        signal = '⚠️数据不足'
        scores['momentum'] = momentum_score
        scores['signal'] = signal
        scores['d2v'] = None
        scores['d2p'] = None

    # 综合评分
    total = (scores['valuation'] * WEIGHTS['valuation'] +
             scores['policy'] * WEIGHTS['policy'] +
             scores['rpe'] * WEIGHTS['rpe'] +
             scores['trend'] * WEIGHTS['trend'] +
             scores['momentum'] * WEIGHTS['momentum'])
    scores['total'] = round(total, 2)

    if total >= 80:
        grade = 'A'
    elif total >= 70:
        grade = 'B'
    elif total >= 60:
        grade = 'C'
    else:
        grade = 'D'
    scores['grade'] = grade

    scores['price'] = price
    scores['volume'] = volume
    scores['pe_ttm'] = pe_ttm if pe_ttm else None
    scores['sector_pe'] = sector_pe
    scores['type'] = stype
    scores['level'] = slevel
    scores['sector'] = sector
    scores['status'] = stock_info.get('status', '')
    scores['name'] = stock_info['name']
    scores['code'] = code

    return scores


# ==================== 主程序 ====================

def main():
    print("=" * 60)
    print(f"📊 股票量化评分系统 V3.1（混合数据源版）")
    print(f"⏰ 分析时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 登录baostock
    lg = bs.login()
    if lg.error_code != '0':
        print(f"❌ baostock登录失败：{lg.error_msg}")
        return

    results = []
    total = len(STOCK_POOL)
    pe_cache = {}

    for idx, stock in enumerate(STOCK_POOL, 1):
        code = stock['code']
        name = stock['name']
        print(f"\n[{idx}/{total}] 处理 {code} {name} ...")

        # 1. 使用baostock获取历史K线
        hist = get_hist_data_baostock(code)
        if hist is None or len(hist) == 0:
            print(f"  ⚠️ 无法获取历史K线，跳过趋势和量价分析")
        else:
            print(f"  📋 历史K线获取成功，共{len(hist)}条")

        # 2. 使用雪球获取实时行情
        quote = get_quote_data_xueqiu(code)
        if quote is None or quote['price'] == 0:
            print(f"  ❌ 无法获取行情数据，跳过该股")
            continue

        print(f"  📋 最新价: {quote['price']}")

        # 3. 计算评分（PE在calc_scores内部通过人机交互获取）
        scores = calc_scores(code, hist, quote, stock['sector'], stock, pe_cache)
        if scores is None:
            print(f"  ❌ 评分计算失败")
            continue

        result = {
            '代码': scores['code'],
            '名称': scores['name'],
            '状态': scores['status'],
            '类型': scores['type'],
            '政策层级': scores['level'],
            '板块': scores['sector'],
            '收盘价': scores['price'],
            '成交量(手)': scores['volume'],
            'TTM PE': scores['pe_ttm'],
            '板块PE': scores['sector_pe'],
            'RPE值': scores['rpe_value'],
            '估值分': scores['valuation'],
            '政策分': scores['policy'],
            'RPE分': scores['rpe'],
            '趋势分': scores['trend'],
            '量价分': scores['momentum'],
            'd2v': scores.get('d2v'),
            'd2p': scores.get('d2p'),
            '双二阶信号': scores.get('signal', ''),
            '综合得分': scores['total'],
            '评级': scores['grade'],
        }
        results.append(result)
        print(f"  ✅ 完成，综合得分：{scores['total']}，评级：{scores['grade']}")

    # 登出baostock
    bs.logout()

    if results:
        df = pd.DataFrame(results)
        df = df.sort_values('综合得分', ascending=False)
        today = datetime.now().strftime('%Y%m%d')
        filename = f"{OUTPUT_PREFIX}_{today}.xlsx"
        filepath = os.path.join(script_dir, filename)

        df.to_excel(filepath, index=False, sheet_name='评分报告')

        print("\n" + "=" * 60)
        print(f"✅ 分析完成！")
        print(f"📁 文件位置：{filepath}")
        print(f"📊 共分析 {len(results)} 只标的")
        print("\n评级分布：")
        grade_order = ['A', 'B', 'C', 'D']
        for grade in grade_order:
            count = len(df[df['评级'] == grade])
            if count > 0:
                print(f"  {grade}级：{count} 只")
        print("=" * 60)
    else:
        print("\n❌ 无有效数据")


if __name__ == "__main__":
    main()