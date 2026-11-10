import math
from scipy.stats import norm

# Các mốc chuẩn trong Bảng A1, A2, A3 của CLSI EP26-Ed2
SR_SWRL_GRID = [1.00, 0.95, 0.90, 0.85, 0.80, 0.75, 0.70, 0.60, 0.50, 0.40, 0.30]
K_OPTIONS = [0.90, 0.80, 0.70, 0.60, 0.55]

ALPHA_BY_LEVELS = {
    1: 0.0500,  # Bảng A1
    2: 0.0250,  # Bảng A2
    3: 0.0167   # Bảng A3
}

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
    """
    Trả về danh sách các phương án (k, N, Power, RL) thỏa mãn desired_power,
    được sắp xếp tối ưu theo Số mẫu ít nhất và Power achieve cao nhất.
    """
    if swrl_pct <= 0 or sr_pct <= 0 or tea_pct <= 0:
        return [{
            "cd_pct": 0.0,
            "cd_swrl": 0.0,
            "sr_swrl": 0.0,
            "power_achieve": 0.0,
            "n_samples": 0,
            "rejection_limit": 0.0,
            "k_factor": 0.0,
            "meets_power": False,
            "note": "Thông số đầu vào phải > 0"
        }]

    cd_pct = tea_pct - 2.0 * swrl_pct
    cd_swrl = cd_pct / swrl_pct
    sr_swrl = sr_pct / swrl_pct

    if cd_pct <= 0:
        return [{
            "cd_pct": round(cd_pct, 3),
            "cd_swrl": round(cd_swrl, 3),
            "sr_swrl": round(sr_swrl, 3),
            "power_achieve": 0.0,
            "n_samples": 0,
            "rejection_limit": round(0.80 * cd_pct, 2),
            "k_factor": 0.80,
            "meets_power": False,
            "note": "CD% <= 0 (CVwrl quá lớn so với TEa)"
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
                "rejection_limit": round(k_cand * cd_pct, 2),
                "k_factor": round(k_cand, 2),
                "meets_power": meets,
                "note": (
                    f"Đạt chuẩn (Mốc tra bảng: CD/Swrl={cd_swrl_tab}, Sr/Swrl={sr_swrl_tab})"
                    if meets else f"Power < {desired_power:.2f}"
                )
            })

    if not valid_options:
        return [{
            "cd_pct": round(cd_pct, 3),
            "cd_swrl": round(cd_swrl, 3),
            "sr_swrl": round(sr_swrl, 3),
            "power_achieve": 0.0,
            "n_samples": 0,
            "rejection_limit": round(0.70 * cd_pct, 2),
            "k_factor": 0.70,
            "meets_power": False,
            "note": "Không tra được trong bảng với M=1 (Cần tăng CD hoặc chạy nhiều lần)"
        }]

    # Lọc các phương án đạt Power mong muốn (nếu không có phương án nào đạt thì giữ lại toàn bộ để tham khảo)
    matching = [opt for opt in valid_options if opt["meets_power"]]
    pool = matching if matching else valid_options

    # Sắp xếp danh sách phù hợp:
    # 1. "min_n_max_power": Ưu tiên Số mẫu N ít nhất -> nếu cùng N thì lấy Power achieve cao nhất
    # 2. "max_power_min_n": Ưu tiên Power achieve cao nhất -> nếu cùng Power thì lấy N ít nhất
    if sort_mode == "max_power_min_n":
        pool.sort(key=lambda x: (-x["power_achieve"], x["n_samples"]))
    else:
        pool.sort(key=lambda x: (x["n_samples"], -x["power_achieve"]))

    return pool