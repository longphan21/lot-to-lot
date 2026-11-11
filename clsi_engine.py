import math
import pandas as pd
from scipy.stats import norm

SR_SWRL_GRID = [1.00, 0.95, 0.90, 0.85, 0.80, 0.75, 0.70, 0.60, 0.50, 0.40, 0.30]
K_OPTIONS = [0.90, 0.80, 0.70, 0.60, 0.55]

ALPHA_BY_LEVELS = {
    1: 0.0500,  # Bảng A1
    2: 0.0250,  # Bảng A2
    3: 0.0167  # Bảng A3
}

MAX_CD_SWRL_BY_LEVELS = {
    1: 5.5,
    2: 6.0,
    3: 6.5
}


def round_down_cd_swrl(val: float, num_levels: int) -> float:
    max_cap = MAX_CD_SWRL_BY_LEVELS.get(num_levels, 6.0)
    if val >= max_cap:
        return max_cap
    if val < 1.0:
        return 1.0
    return math.floor(val * 2.0) / 2.0


def round_down_sr_swrl(val: float) -> float:
    if val >= 1.00:
        return 1.00
    for grid_val in SR_SWRL_GRID:
        if val >= grid_val - 1e-9:
            return grid_val
    return 0.30


def calc_n_and_power(cd_swrl_tab: float, sr_swrl_tab: float, k: float, num_levels: int):
    if num_levels in ALPHA_BY_LEVELS:
        alpha = ALPHA_BY_LEVELS[num_levels]
    else:
        alpha = 1.0 - (1.0 - 0.05) ** (1.0 / max(1, num_levels))

    z_crit = norm.ppf(1.0 - alpha / 2.0)

    denom = 0.5 * ((k * cd_swrl_tab) / z_crit) ** 2 + (sr_swrl_tab ** 2) - 1.0
    if denom <= 0:
        return None, None

    n_raw = (sr_swrl_tab ** 2) / denom
    if n_raw <= 0:
        return None, None

    n_samples = max(1, math.ceil(n_raw))

    spread = math.sqrt(2.0) * math.sqrt(((1.0 / n_samples) - 1.0) * (sr_swrl_tab ** 2) + 1.0)
    term1 = ((k - 1.0) * cd_swrl_tab) / spread
    term2 = ((-k - 1.0) * cd_swrl_tab) / spread
    beta = norm.cdf(term1) - norm.cdf(term2)
    power = 1.0 - beta

    return int(n_samples), round(float(power), 3)


def get_candidate_options_for_row(
        tea_pct: float,
        swrl_pct: float,
        sr_pct: float,
        num_levels: int = 2,
        desired_power: float = 0.80,
        sort_mode: str = "min_n_max_power"
) -> list:
    if swrl_pct <= 0 or sr_pct <= 0 or tea_pct <= 0:
        return [{
            "cd_pct": 0.0, "cd_swrl": 0.0, "sr_swrl": 0.0,
            "power_achieve": 0.0, "n_samples": 0, "rejection_limit": 0.0,
            "k_factor": 0.0, "meets_power": False, "note": "Thông số đầu vào phải > 0"
        }]

    cd_pct = tea_pct - 2.0 * swrl_pct
    cd_swrl = cd_pct / swrl_pct
    sr_swrl = sr_pct / swrl_pct

    if cd_pct <= 0:
        return [{
            "cd_pct": round(cd_pct, 3), "cd_swrl": round(cd_swrl, 3), "sr_swrl": round(sr_swrl, 3),
            "power_achieve": 0.0, "n_samples": 0, "rejection_limit": round(0.80 * cd_pct, 2),
            "k_factor": 0.80, "meets_power": False, "note": "CD% <= 0 (CVwrl quá lớn so với TEa)"
        }]

    cd_swrl_tab = round_down_cd_swrl(cd_swrl, num_levels)
    sr_swrl_tab = round_down_sr_swrl(sr_swrl)

    valid_options = []
    for k_cand in K_OPTIONS:
        n_val, p_val = calc_n_and_power(cd_swrl_tab, sr_swrl_tab, k_cand, num_levels)
        if n_val is not None:
            meets = p_val >= desired_power
            valid_options.append({
                "cd_pct": round(cd_pct, 3),
                "cd_swrl": round(cd_swrl, 3),
                "sr_swrl": round(sr_swrl, 3),
                "power_achieve": round(p_val, 3),
                "n_samples": int(n_val),
                "rejection_limit": round(k_cand * cd_pct, 3),
                "k_factor": round(k_cand, 2),
                "meets_power": meets,
                "note": (
                    f"Đạt chuẩn (Tra bảng: CD/Swrl={cd_swrl_tab}, Sr/Swrl={sr_swrl_tab})"
                    if meets else f"Power < {desired_power:.2f}"
                )
            })

    if not valid_options:
        return [{
            "cd_pct": round(cd_pct, 3), "cd_swrl": round(cd_swrl, 3), "sr_swrl": round(sr_swrl, 3),
            "power_achieve": 0.0, "n_samples": 0, "rejection_limit": round(0.70 * cd_pct, 3),
            "k_factor": 0.70, "meets_power": False,
            "note": "Không tra được trong bảng với M=1 (Cần tăng CD hoặc chạy nhiều lần)"
        }]

    matching = [opt for opt in valid_options if opt["meets_power"]]
    pool = matching if matching else valid_options

    if sort_mode == "max_power_min_n":
        pool.sort(key=lambda x: (-x["power_achieve"], x["n_samples"]))
    else:
        pool.sort(key=lambda x: (x["n_samples"], -x["power_achieve"]))

    return pool


def evaluate_stage2_level(
        samples_df: pd.DataFrame,
        rl_pct: float
) -> dict:
    """
    Tính toán kết quả Stage 2 khớp 100% với hàm AVERAGE của khoa:
    1. d_i = |LOT_new - LOT_old| (Lấy trị tuyệt đối từng mẫu trước)
    2. mean_d = AVERAGE(d_i)
    3. mean_old = AVERAGE(LOT_old)
    4. mean_d_pct = mean_d * 100 / mean_old
    """
    valid_rows = []
    for idx, row in samples_df.iterrows():
        val_old = row.get("Kết quả LOT cũ")
        val_new = row.get("Kết quả LOT mới")
        if pd.notna(val_old) and pd.notna(val_new):
            try:
                f_old = float(val_old)
                f_new = float(val_new)
                if f_old > 0 and f_new > 0:
                    sid = str(row.get("SID (Mã mẫu)", "")).strip() or f"Mẫu #{len(valid_rows) + 1}"

                    # Chốt công thức trị tuyệt đối để khớp Excel: =AVERAGE(L2:L3)
                    d_abs = abs(f_new - f_old)

                    valid_rows.append({
                        "STT": len(valid_rows) + 1,
                        "SID": sid,
                        "LOT_old_val": f_old,
                        "LOT_new_val": f_new,
                        "d": round(d_abs, 4)
                    })
            except (ValueError, TypeError):
                continue

    if not valid_rows:
        return {
            "has_data": False,
            "n_actual": 0,
            "mean_old": 0.0,
            "mean_new": 0.0,
            "mean_d": 0.0,
            "mean_d_pct": 0.0,
            "rl_pct": round(rl_pct, 3),
            "result": "CHƯA NHẬP MẪU",
            "details": pd.DataFrame()
        }

    df_valid = pd.DataFrame(valid_rows)
    n_actual = len(df_valid)

    mean_old = float(df_valid["LOT_old_val"].mean())
    mean_new = float(df_valid["LOT_new_val"].mean())
    mean_d = float(df_valid["d"].mean())

    # Khớp công thức: =AVERAGE(L2:L3)*100/AVERAGE(H2:H3)
    mean_d_pct = (mean_d * 100.0 / mean_old) if mean_old > 0 else 0.0

    is_pass = mean_d_pct < rl_pct

    return {
        "has_data": True,
        "n_actual": n_actual,
        "mean_old": round(mean_old, 4),
        "mean_new": round(mean_new, 4),
        "mean_d": round(mean_d, 6),
        "mean_d_pct": round(mean_d_pct, 3),
        "rl_pct": round(rl_pct, 3),
        "result": "PASS" if is_pass else "FAIL",
        "details": df_valid
    }