import math
from scipy.stats import norm

# Các mốc chuẩn trong Bảng A1, A2, A3 của CLSI EP26-Ed2
SR_SWRL_GRID = [1.00, 0.95, 0.90, 0.85, 0.80, 0.75, 0.70, 0.60, 0.50, 0.40, 0.30]
K_OPTIONS = [0.90, 0.80, 0.70, 0.60, 0.55]

# Giá trị alpha theo Phụ lục B (Mục B3) cho L = 1, 2, 3 mức nồng độ
ALPHA_BY_LEVELS = {
    1: 0.0500,  # Bảng A1
    2: 0.0250,  # Bảng A2
    3: 0.0167   # Bảng A3
}

# Ngưỡng tối đa của tỷ số CD/Swrl trong Bảng A1 (5.5), A2 (6.0), A3 (6.5)
MAX_CD_SWRL_BY_LEVELS = {
    1: 5.5,
    2: 6.0,
    3: 6.5
}


def round_down_cd_swrl(val: float, num_levels: int) -> float:
    """Làm tròn xuống CD/Swrl theo bước nhảy 0.5 của bảng CLSI EP26."""
    max_cap = MAX_CD_SWRL_BY_LEVELS.get(num_levels, 6.0)
    if val >= max_cap:
        return max_cap
    if val < 1.0:
        return 1.0
    return math.floor(val * 2.0) / 2.0


def round_down_sr_swrl(val: float) -> float:
    """Làm tròn xuống Sr/Swrl về mốc gần nhất trong bảng CLSI EP26."""
    if val >= 1.00:
        return 1.00
    for grid_val in SR_SWRL_GRID:
        if val >= grid_val - 1e-9:
            return grid_val
    return 0.30


def calc_n_and_power(cd_swrl_tab: float, sr_swrl_tab: float, k: float, num_levels: int):
    """
    Tính số mẫu N và Power đạt được theo công thức (B23), (B27), (B28)
    trong Phụ lục B - CLSI EP26-Ed2 với số lần chạy M = 1.
    """
    if num_levels in ALPHA_BY_LEVELS:
        alpha = ALPHA_BY_LEVELS[num_levels]
    else:
        alpha = 1.0 - (1.0 - 0.05) ** (1.0 / max(1, num_levels))

    z_crit = norm.ppf(1.0 - alpha / 2.0)

    # Công thức B23 (với M = 1)
    denom = 0.5 * ((k * cd_swrl_tab) / z_crit) ** 2 + (sr_swrl_tab ** 2) - 1.0
    if denom <= 0:
        return None, None

    n_raw = (sr_swrl_tab ** 2) / denom
    if n_raw <= 0:
        return None, None

    n_samples = max(1, math.ceil(n_raw))

    # Công thức B27 & B28 tính Beta và Power = 1 - Beta
    spread = math.sqrt(2.0) * math.sqrt(((1.0 / n_samples) - 1.0) * (sr_swrl_tab ** 2) + 1.0)
    term1 = ((k - 1.0) * cd_swrl_tab) / spread
    term2 = ((-k - 1.0) * cd_swrl_tab) / spread
    beta = norm.cdf(term1) - norm.cdf(term2)
    power = 1.0 - beta

    return int(n_samples), round(float(power), 3)


def evaluate_stage1_row(
    tea_pct: float,
    swrl_pct: float,
    sr_pct: float,
    num_levels: int = 2,
    target_power: float = 0.80,
    manual_k: float = None
) -> dict:
    """
    Tính toán toàn bộ thông số cho 1 dòng nồng độ giống bảng Excel:
    - CD% = TEa% - 2 * Swrl%
    - CD/Swrl = CD% / Swrl%
    - Sr/Swrl = Sr% / Swrl%
    - Tra cứu Power achieve, Number of sample (N), hệ số k và Rejection limit (RL%).
    """
    if swrl_pct <= 0 or sr_pct <= 0 or tea_pct <= 0:
        return {
            "cd_pct": 0.0,
            "cd_swrl": 0.0,
            "sr_swrl": 0.0,
            "power_achieve": 0.0,
            "n_samples": 0,
            "rejection_limit": 0.0,
            "k_factor": 0.0,
            "note": "Thông số đầu vào phải > 0"
        }

    cd_pct = tea_pct - 2.0 * swrl_pct
    cd_swrl = cd_pct / swrl_pct
    sr_swrl = sr_pct / swrl_pct

    # Nếu CD% <= 0 (như mức 0.188 của máy Roche khi Swrl = 10% > TEa/2)
    if cd_pct <= 0:
        k_used = manual_k if (manual_k and manual_k > 0) else 0.80
        return {
            "cd_pct": round(cd_pct, 3),
            "cd_swrl": round(cd_swrl, 3),
            "sr_swrl": round(sr_swrl, 3),
            "power_achieve": 0.0,
            "n_samples": 0,
            "rejection_limit": round(k_used * cd_pct, 2),
            "k_factor": round(k_used, 2),
            "note": "CD% <= 0 (CVwrl quá lớn so với TEa)"
        }

    # Làm tròn xuống theo đúng hướng dẫn tra Bảng A1-A3 của CLSI EP26
    cd_swrl_tab = round_down_cd_swrl(cd_swrl, num_levels)
    sr_swrl_tab = round_down_sr_swrl(sr_swrl)

    chosen_k = None
    chosen_n = None
    chosen_power = None
    note = "Đạt chuẩn CLSI EP26"

    if manual_k is not None and manual_k > 0:
        # Người dùng chỉ định sẵn hệ số k (ví dụ 0.7 hoặc 0.8)
        n_val, p_val = calc_n_and_power(cd_swrl_tab, sr_swrl_tab, manual_k, num_levels)
        chosen_k = manual_k
        chosen_n = n_val if n_val is not None else 0
        chosen_power = p_val if p_val is not None else 0.0
        if n_val is None:
            note = "Không đạt α <= 5% với M=1 ở hệ số k này"
    else:
        # Tự động quét k từ 0.90 xuống 0.55 để tìm mốc đầu tiên đạt Power >= target_power
        for k_cand in K_OPTIONS:
            n_val, p_val = calc_n_and_power(cd_swrl_tab, sr_swrl_tab, k_cand, num_levels)
            if n_val is not None and p_val >= target_power:
                chosen_k = k_cand
                chosen_n = n_val
                chosen_power = p_val
                break

        # Nếu quét hết vẫn chưa đạt target_power thì lấy mốc k nhỏ nhất có thể tính được
        if chosen_k is None:
            for k_cand in reversed(K_OPTIONS):
                n_val, p_val = calc_n_and_power(cd_swrl_tab, sr_swrl_tab, k_cand, num_levels)
                if n_val is not None:
                    chosen_k = k_cand
                    chosen_n = n_val
                    chosen_power = p_val
                    note = f"Power tối đa chỉ đạt {p_val:.3f}"
                    break

        if chosen_k is None:
            chosen_k = 0.70
            chosen_n = 0
            chosen_power = 0.0
            note = "Cần tăng CD hoặc tăng số lần chạy (M > 1)"

    rejection_limit = chosen_k * cd_pct

    return {
        "cd_pct": round(cd_pct, 3),
        "cd_swrl": round(cd_swrl, 3),
        "sr_swrl": round(sr_swrl, 3),
        "power_achieve": round(chosen_power, 3),
        "n_samples": int(chosen_n),
        "rejection_limit": round(rejection_limit, 2),
        "k_factor": round(chosen_k, 2),
        "note": note
    }