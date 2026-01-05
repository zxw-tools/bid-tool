# -*- coding: utf-8 -*-
"""
bid_tool2_final.py  （信用分优先版）

新增：
- 从 Excel 导入“信用分”列；
- 中标判定规则：
  1. 先比报价得分（score），高者优先；
  2. 得分相等时，比信用分（credit），高者优先；
  3. 若得分与信用分都相等，则比总报价（bid），低者优先；
  4. 若三项都相同，则视为并列中标。

评分仍然全部基于“评标价”（= 报价 - 暂列金/暂估价）来计算偏差率和得分。
"""

import itertools
from decimal import Decimal, getcontext, ROUND_HALF_UP

import pandas as pd

# ========= Decimal 精度 =========
getcontext().prec = 28  # 足够高，避免浮点误差


def d(x) -> Decimal:
    """统一转成 Decimal"""
    return Decimal(str(x))


def round_decimal(value, places) -> Decimal:
    """四舍五入保留 places 位小数"""
    q = Decimal("1").scaleb(-places)  # = 10^(-places)
    return d(value).quantize(q, rounding=ROUND_HALF_UP)


# ========= 基准价计算：一次 / 二次 / 三次平均 + 去高去低 =========
def compute_baseline(eval_prices, method, remove_low, remove_high):
    """
    eval_prices: 当前组合所有评标价 (Decimal)
    method: 1 = 一次平均, 2 = 二次平均, 3 = 三次平均
    remove_low / remove_high: 去掉的最低 / 最高家数
    """
    prices = [p for p in eval_prices if p is not None]
    prices = sorted(prices)
    n = len(prices)
    if n == 0:
        raise ValueError("没有有效评标价")

    if remove_low + remove_high >= n:
        raise ValueError("去掉最高/最低后无剩余评标价")

    # 先去掉最高/最低
    start = remove_low
    end = n - remove_high if remove_high > 0 else n
    prices = prices[start:end]
    n = len(prices)

    # 一次平均
    avg1 = sum(prices) / d(n)
    if method == 1:
        return round_decimal(avg1, 2)

    # 二次平均：取“小于一次平均”的评标价再平均
    lower = [p for p in prices if p < avg1]
    if not lower:
        lower = prices[:]
    avg2 = sum(lower) / d(len(lower))
    if method == 2:
        return round_decimal(avg2, 2)

    # 三次平均：
    # ① avg1：一次平均
    # ② avg2：二次平均
    # ③ 所有 (avg2 < 评标价 < avg1) 的评标价 + avg1 + avg2 再平均
    band = [p for p in prices if p > avg2 and p < avg1]
    base_set = band + [avg1, avg2] if band else [avg1, avg2]
    avg3 = sum(base_set) / d(len(base_set))
    return round_decimal(avg3, 2)


# ========= 用评标价计算得分 =========
def compute_scores_for_group(eval_prices, bids, baseline, full_score, e1, e2):
    """
    eval_prices: 本组合每家的评标价列表 (Decimal)
    bids: 对应的总报价列表 (Decimal)
    baseline: 评标基准价 (Decimal)
    full_score: 报价满分
    e1/e2: 高于基准价 / 低于等于基准价的系数
    """
    scores = []
    for eval_p, bid in zip(eval_prices, bids):
        # 偏差率 r = (评标价 - B) / B
        r = (eval_p - baseline) / baseline
        # 先四舍五入到 4 位小数
        r4 = round_decimal(r, 4)

        if eval_p > baseline:
            score = full_score - (r4 * d(100)) * e1
        else:
            score = full_score + (r4 * d(100)) * e2

        # 最终得分保留 2 位小数
        scores.append(round_decimal(score, 2))

    return scores


# ========= 统一的“选中标单位”的函数（加入信用分） =========
def pick_winner_positions(scores, comb_bids, comb_credits):
    """
    选出组合中的中标单位位置索引列表（相对本组合的索引）：
    1. 先看报价得分 scores，得分最高者中标；
    2. 得分并列时，看信用分 comb_credits，信用高者中标；
    3. 若得分 & 信用分都相同，则看总报价 comb_bids，总报价低者中标；
    4. 若再完全一样，则这些都算并列中标。
    """
    # ① 得分最高
    max_score = max(scores)
    candidate_pos = [i for i, s in enumerate(scores) if s == max_score]
    if len(candidate_pos) == 1:
        return candidate_pos

    # ② 在得分最高的候选中，信用分最高
    candidate_credits = [comb_credits[i] for i in candidate_pos]
    max_credit = max(candidate_credits)
    credit_best_pos = [i for i in candidate_pos if comb_credits[i] == max_credit]
    if len(credit_best_pos) == 1:
        return credit_best_pos

    # ③ 信用分也并列时，在这些人里比较总报价，报价越低越优
    min_bid = min(comb_bids[i] for i in credit_best_pos)
    final_pos = [i for i in credit_best_pos if comb_bids[i] == min_bid]
    return final_pos


# ========= 输入小工具 =========
def input_int(prompt, allow_empty=False, default=None):
    while True:
        s = input(prompt).strip()
        if not s:
            if allow_empty:
                return default
            print("不能为空，请重新输入。")
            continue
        try:
            return int(s)
        except ValueError:
            print("请输入整数。")


def input_decimal(prompt, allow_empty=False, default=None):
    while True:
        s = input(prompt).strip()
        if not s:
            if allow_empty:
                return default
            print("不能为空，请重新输入。")
            continue
        try:
            return d(s)
        except Exception:
            print("请输入数字。")


def yes_no(prompt, default="y"):
    while True:
        s = input(prompt).strip().lower()
        if not s and default:
            return default == "y"
        if s in ("y", "yes", "是"):
            return True
        if s in ("n", "no", "否"):
            return False
        print("请输入 Y 或 N。")


# ========= 从 Excel 导入（增加信用分） =========
def load_from_excel():
    """
    从 Excel 中自动识别列：

    必选：
    - 公司名称   → ["公司名称", "名称", "单位", "投标人", "投标单位", "供应商"]
    - 总报价     → ["总报价", "报价", "投标报价", "投标价", "报价（元）", "投标总价"]

    可选：
    - 评标价     → ["评标价", "评标报价", "评标价（元）"]
    - 暂列金/暂估价 → ["暂列金+暂估价", "暂列金", "暂估价"]
    - 信用分     → ["信用分", "信用得分", "信用评分", "信用分数", "信用"]

    统一逻辑：
    1）总报价 & 评标价都有：  总报价=总报价；评标价=评标价；
    2）总报价 & 暂列金都有：  评标价 = 总报价 - 暂列金；
    3）只有评标价：            总报价 = 评标价；
    4）只有总报价：            评标价 = 总报价。
    5）信用分列若不存在，则默认信用分=0。
    """
    path = input(
        "请输入 Excel 路径（可拖进来，例如 C:\\\\Users\\\\你\\\\Desktop\\\\报价表.xlsx）：\n"
    ).strip().strip('"')
    try:
        df = pd.read_excel(path)
    except Exception as e:
        print("❌ Excel 读取失败：", e)
        return None, None, None, None

    name_candidates = ["公司名称", "名称", "单位", "投标人", "投标单位", "供应商"]
    total_candidates = ["总报价", "报价", "投标报价", "投标价", "报价（元）", "投标总价"]
    eval_candidates = ["评标价", "评标报价", "评标价（元）"]
    provisional_candidates = ["暂列金+暂估价", "暂列金", "暂估价"]
    credit_candidates = ["信用分", "信用得分", "信用评分", "信用分数", "信用"]

    name_col = total_col = eval_col = provisional_col = credit_col = None

    for c in df.columns:
        if name_col is None and c in name_candidates:
            name_col = c
        if total_col is None and c in total_candidates:
            total_col = c
        if eval_col is None and c in eval_candidates:
            eval_col = c
        if provisional_col is None and c in provisional_candidates:
            provisional_col = c
        if credit_col is None and c in credit_candidates:
            credit_col = c

    if name_col is None:
        print("❌ 未找到“公司名称/单位”等列。表头为：", list(df.columns))
        return None, None, None, None

    if total_col is None and eval_col is None:
        print("❌ 未找到“报价/评标价”等列。表头为：", list(df.columns))
        return None, None, None, None

    companies = []
    bids = []
    eval_prices = []
    credits = []

    for _, row in df.iterrows():
        name = row.get(name_col, "")
        if pd.isna(name):
            continue
        name = str(name).strip()
        if not name:
            continue

        total_val = row.get(total_col) if total_col is not None else None
        eval_val = row.get(eval_col) if eval_col is not None else None
        prov_val = row.get(provisional_col) if provisional_col is not None else None
        cred_val = row.get(credit_col) if credit_col is not None else None

        t = d(total_val) if total_val not in (None, "") and not pd.isna(total_val) else None
        e = d(eval_val) if eval_val not in (None, "") and not pd.isna(eval_val) else None
        p = d(prov_val) if prov_val not in (None, "") and not pd.isna(prov_val) else None

        if cred_val in (None, "") or pd.isna(cred_val):
            c = d(0)
        else:
            c = d(cred_val)

        # 统一生成总报价 / 评标价
        if t is not None and e is not None:
            total = t
            eval_p = e
        elif t is not None and p is not None:
            total = t
            eval_p = t - p
        elif e is not None:
            total = e
            eval_p = e
        elif t is not None:
            total = t
            eval_p = t
        else:
            continue

        companies.append(name)
        bids.append(total)
        eval_prices.append(eval_p)
        credits.append(c)

    if not companies:
        print("❌ 没有解析到任何有效公司与报价，请检查表格。")
        return None, None, None, None

    print(f"\n✅ 已从 Excel 读入 {len(companies)} 家单位：")
    for i, (n, b, e, c) in enumerate(zip(companies, bids, eval_prices, credits), start=1):
        print(f"{i}. {n}  总报价={b:,.2f}  评标价={e:,.2f}  信用分={c}")

    return companies, bids, eval_prices, credits


# ========= 功能 1：中标概率 =========
def do_calc_probabilities(
    N,
    K,
    companies,
    bids,
    eval_prices_all,
    credits_all,
    limit_price,
    full_score,
    e1,
    e2,
    method,
    remove_low,
    remove_high,
):
    print("\n>>> 正在枚举所有组合，按“得分 + 信用 + 报价”真实规则计算中标概率...")

    # 限价过滤（按总报价判断）
    valid_indices = []
    for idx in range(N):
        if limit_price is None or bids[idx] <= limit_price:
            valid_indices.append(idx)

    if len(valid_indices) < K:
        print("⚠ 有效投标单位不足 K 家，结果仅供参考。")
        return

    from decimal import Decimal as D

    win_counts = [D("0")] * N
    total_subsets = D("0")

    for comb in itertools.combinations(valid_indices, K):
        indices = list(comb)
        eval_prices = [eval_prices_all[i] for i in indices]
        comb_bids = [bids[i] for i in indices]
        comb_credits = [credits_all[i] for i in indices]

        try:
            B = compute_baseline(eval_prices, method, remove_low, remove_high)
        except ValueError:
            continue

        scores = compute_scores_for_group(eval_prices, comb_bids, B, full_score, e1, e2)
        winner_pos = pick_winner_positions(scores, comb_bids, comb_credits)

        total_subsets += D("1")
        for j in winner_pos:
            real_idx = indices[j]
            win_counts[real_idx] += D("1")

    print("\n=== 中标概率结果（组合个数逻辑） ===")
    for idx in range(N):
        name = companies[idx]
        bid = bids[idx]
        cred = credits_all[idx]
        wins = win_counts[idx]
        if total_subsets == 0:
            prob = D("0")
        else:
            prob = wins / total_subsets
        print(
            f"{idx+1}. {name}: 中标次数 = {wins:.0f}，"
            f"中标概率 = {prob * 100:.2f}%（总报价 = {bid:,.2f}，信用分 = {cred}）"
        )


# ========= 功能 2：手动指定一组公司 =========
def do_calc_manual_group(
    N,
    companies,
    bids,
    eval_prices_all,
    credits_all,
    limit_price,
    full_score,
    e1,
    e2,
    method,
    remove_low,
    remove_high,
):
    K = input_int("\n本次参与评审的公司数量 K（通常 = 5）：")
    if K <= 0 or K > N:
        print("K 必须在 1~N 之间。")
        return

    print("\n当前公司：")
    for i in range(N):
        print(
            f"{i+1}. {companies[i]}  总报价 = {bids[i]:,.2f}  "
            f"评标价 = {eval_prices_all[i]:,.2f}  信用分 = {credits_all[i]}"
        )

    s = input(
        f"\n请输入本次参与评审的 {K} 家公司序号（空格分隔，例如：2 3 6 7 10）：\n"
    ).strip()
    try:
        indices = [int(x) - 1 for x in s.split()]
    except ValueError:
        print("❌ 输入格式错误，请只输入数字序号，用空格分隔。")
        return

    if len(indices) != K:
        print(f"❌ 你一共输入了 {len(indices)} 家公司，但 K = {K}。")
        return

    if len(set(indices)) != len(indices):
        print("❌ 序号有重复，请重新输入。")
        return

    for idx in indices:
        if not (0 <= idx < N):
            print("❌ 存在超出范围的序号，请检查。")
            return

    if limit_price is not None:
        for idx in indices:
            if bids[idx] > limit_price:
                print(
                    f"❌ {companies[idx]} 的总报价 {bids[idx]:,.2f} 超过限价 {limit_price:,.2f}，不能参与本次评审。"
                )
                return

    eval_prices = [eval_prices_all[i] for i in indices]
    comb_bids = [bids[i] for i in indices]
    comb_credits = [credits_all[i] for i in indices]

    try:
        B = compute_baseline(eval_prices, method, remove_low, remove_high)
    except ValueError as e:
        print("❌ 无法计算基准价：", e)
        return

    scores = compute_scores_for_group(eval_prices, comb_bids, B, full_score, e1, e2)
    winner_pos = pick_winner_positions(scores, comb_bids, comb_credits)

    print(f"\n本组合评标基准价 B = {B:,.2f}")
    print("本组合各家公司报价得分（按得分高->低；得分相同则信用分高->低；若再相同则报价低->高）：")

    rows = []
    for pos, (idx, score) in enumerate(zip(indices, scores)):
        name = companies[idx]
        bid = comb_bids[pos]
        eval_p = eval_prices[pos]
        cred = comb_credits[pos]
        is_winner = pos in winner_pos
        rows.append((score, cred, bid, name, idx, eval_p, is_winner))

    # 排序：得分降序 -> 信用分降序 -> 报价升序
    rows.sort(key=lambda x: (-x[0], -x[1], x[2]))

    for score, cred, bid, name, idx, eval_p, is_winner in rows:
        flag = "<== 中标" if is_winner else ""
        print(
            f"{idx+1}. {name}: 总报价 = {bid:,.2f}，评标价 = {eval_p:,.2f}，"
            f"信用分 = {cred}，得分 = {score:.2f} {flag}"
        )


# ========= 功能 3：指定公司，列出所有“真正能中标”的组合 =========
def do_list_winning_combos(
    N,
    companies,
    bids,
    eval_prices_all,
    credits_all,
    limit_price,
    full_score,
    e1,
    e2,
    method,
    remove_low,
    remove_high,
):
    from decimal import Decimal as D

    if limit_price is not None:
        valid_indices = [i for i, b in enumerate(bids) if b <= limit_price]
        if not valid_indices:
            print("❌ 没有任何报价在限价以内，无法组合。")
            return
    else:
        valid_indices = list(range(N))

    K = input_int("\n每组参与评审的公司数量 K（通常 = 5）：")
    if K <= 0 or K > len(valid_indices):
        print("K 必须在 1~有效公司数 之间。")
        return

    target = input_int(f"请输入要分析的公司序号（1~{N}）：")
    target_idx = target - 1
    if target_idx not in valid_indices:
        print("该公司报价超过限价或不存在。")
        return

    win_combos = []
    print("\n>>> 正在穷举组合，按“得分 + 信用分 + 报价”规则筛选能让该公司中标的组合...")

    for comb in itertools.combinations(valid_indices, K):
        indices = list(comb)
        if target_idx not in indices:
            continue

        eval_prices = [eval_prices_all[i] for i in indices]
        comb_bids = [bids[i] for i in indices]
        comb_credits = [credits_all[i] for i in indices]

        try:
            B = compute_baseline(eval_prices, method, remove_low, remove_high)
        except ValueError:
            continue

        scores = compute_scores_for_group(eval_prices, comb_bids, B, full_score, e1, e2)
        winner_pos = pick_winner_positions(scores, comb_bids, comb_credits)

        local_pos = indices.index(target_idx)
        if local_pos in winner_pos:
            win_combos.append([i + 1 for i in indices])

    if not win_combos:
        print(f"公司“{companies[target_idx]}”在任何组合中都无法成为最终中标单位。")
        return

    print(f"\n共找到 {len(win_combos)} 种组合，可以让“{companies[target_idx]}”成为最终中标单位：")
    for combo in win_combos:
        print(combo)


# ========= 主程序 =========
def main():
    print("========== 投标报价模拟工具（bid_tool2_final，信用分优先版） ==========")
    print("中标规则：先比报价得分；得分相等时先比信用分；若信用也相等，再比总报价（报价低者优先）。\n")

    use_excel = yes_no("是否从 Excel 导入公司名称、报价和信用分？(Y/N，回车默认 Y)：", default="y")

    companies = []
    bids = []
    eval_prices_all = []
    credits_all = []

    if use_excel:
        companies, bids, eval_prices_all, credits_all = load_from_excel()
        if companies is None:
            print("❌ Excel 导入失败，程序结束。")
            return
        N = len(companies)
        print(f"\n当前共有 {N} 家投标单位。")
    else:
        N = input_int("投标单位总数 N：")
        companies = []
        bids = []
        eval_prices_all = []
        credits_all = []

        provisional_sum = input_decimal(
            "统一的暂列金+暂估价（如无则填 0）：", allow_empty=True, default=d(0)
        )

        for i in range(N):
            name = input(f"请输入第 {i+1} 家单位名称：").strip()
            if not name:
                name = f"公司{i+1}"
            companies.append(name)

            bid_val = input_decimal(f"请输入 {name} 的总报价（含暂列金）: ")
            bids.append(bid_val)

            eval_val = bid_val - provisional_sum
            eval_prices_all.append(eval_val)

            cred_val = input_decimal(
                f"请输入 {name} 的信用分（如无则填 0）: ",
                allow_empty=True,
                default=d(0),
            )
            credits_all.append(cred_val)

    K_prob = input_int("计算中标概率时，每次抽取的单位数 K（例如 5）：")
    limit_price = input_decimal("限价（可为空，直接回车则不限制）：", allow_empty=True, default=None)
    full_score = input_decimal(
        "报价满分（默认 100，可直接回车）：", allow_empty=True, default=d(100)
    )
    e1 = input_decimal("E1（高于基准价的扣分系数，例如 0.6）：")
    e2 = input_decimal("E2（低于或等于基准价的加减系数，例如 0.3）：")

    remove_low = input_int("计算基准价时去掉的最低家数（如 0）：", allow_empty=True, default=0)
    remove_high = input_int("计算基准价时去掉的最高家数（如 0）：", allow_empty=True, default=0)

    print("\n基准价计算方式：")
    print("  1. 一次平均")
    print("  2. 二次平均")
    print("  3. 三次平均")
    method = input_int("请选择（1/2/3，回车默认 1）：", allow_empty=True, default=1)
    if method not in (1, 2, 3):
        print("无效输入，按 1（一次平均）处理。")
        method = 1

    while True:
        print("\n===== 功能菜单 =====")
        print("1. 统计每家公司中标概率（按“得分→信用→报价”真实中标规则）")
        print("2. 手动指定一组公司，计算本组合基准价 + 报价得分 + 中标单位")
        print("3. 指定某家公司，列出所有能让它“真正中标”的序号组合")
        print("4. 重新设置参数 / 重新导入 Excel / 换一组数据")
        print("5. 退出")
        choice = input("请输入序号：").strip()

        if choice == "1":
            do_calc_probabilities(
                N,
                K_prob,
                companies,
                bids,
                eval_prices_all,
                credits_all,
                limit_price,
                full_score,
                e1,
                e2,
                method,
                remove_low,
                remove_high,
            )
        elif choice == "2":
            do_calc_manual_group(
                N,
                companies,
                bids,
                eval_prices_all,
                credits_all,
                limit_price,
                full_score,
                e1,
                e2,
                method,
                remove_low,
                remove_high,
            )
        elif choice == "3":
            do_list_winning_combos(
                N,
                companies,
                bids,
                eval_prices_all,
                credits_all,
                limit_price,
                full_score,
                e1,
                e2,
                method,
                remove_low,
                remove_high,
            )
        elif choice == "4":
            print("\n>>> 重新开始整个流程...\n")
            return main()
        elif choice == "5":
            print("程序结束，再见！")
            break
        else:
            print("无效选择，请重新输入。")


if __name__ == "__main__":
    main()
