"""
Thong so van hanh may (TSVH) - tra theo to (cong doan) + kho/kich thuoc.

Du lieu duoc trich xuat san tu 5 file TSVH*.xlsx trong thu muc lam viec, dung dang
dict tinh de tra nhanh thay vi mo lai file .xlsx moi lan chay 1 don hang. Neu cong
ty cap nhat TSVH (doi may, doi thong so chuan...), phai cap nhat lai cac bang duoi
day cho khop - khong tu dong doc lai tu file goc.

Nguon: TSVH TRANG.xlsx, TSVH DAN CAT.xlsx, TSVH THOI.xlsx (doc thu cong ngay
2026-07-30, xem chi tiet trong references/cong-thuc-dinh-muc.md).
"""

# ---------------------------------------------------------------------------
# TRANG - "Toc do may" la quy tac CO DINH, KHONG tra theo TSVH (theo xac nhan
# cua nguoi dung ngay 2026-07-30: xuong luon chinh may ve muc nay bat ke TSVH
# ghi gi).
TRANG_TOC_DO_MAY_CO_DINH = "97 ± 5"

# Toc do dun keo (Hz) theo kho manh (mm) - giong nhau cho ca Bao BOPP va Bao KP.
TRANG_TOC_DO_DUN_KEO = {
    820: "32 ± 1", 860: "43 ± 1", 960: "53 ± 1", 1060: "63 ± 1",
    1160: "73 ± 1", 1220: "83 ± 1", 1260: "83 ± 1",
}

# CHU Y: "Chieu rong khuon xa keo" trong TSVH TRANG.xlsx co bang tra rieng theo
# kho manh (VD K1160 -> 220±20mm), NHUNG theo xac nhan cua nguoi dung ngay
# 2026-07-30, o day luon dung dung "Chieu rong khuon xa keo = Kho manh" (vd
# 1160mm), KHONG dung bang tra do trong TSVH. Giu lai bang goc o day chi de
# tham khao/doi chieu neu sau nay can xac nhan lai.
TRANG_CHIEU_RONG_KHUON_XA_KEO_TSVH_THAM_KHAO_KP = {
    860: "70 ± 20", 960: "120 ± 20", 1060: "170 ± 20",
    1160: "220 ± 20", 1220: "270 ± 20", 1260: "320 ± 20",
}

import unicodedata  # dung chung cho ca DAN_CAT_KHACH_DAC_BIET va bang TSVH chung


def trang_tra_toc_do_dun_keo(kho_manh_mm):
    kho = int(round(kho_manh_mm))
    if kho in TRANG_TOC_DO_DUN_KEO:
        return TRANG_TOC_DO_DUN_KEO[kho]
    return None  # khong khop -> de trong, to vang, hoi lai


def trang_chieu_rong_khuon_xa_keo(kho_manh_mm):
    """Theo quy tac da xac nhan: = kho manh (mm), KHONG tra bang TSVH."""
    return int(round(kho_manh_mm))


# ---------------------------------------------------------------------------
# DAN CAT - co muc rieng theo khach hang dac biet (uu tien tra o day truoc),
# neu khong co thi dung muc chung "GENERIC".
#
# Moi muc: dict voi key la (chieu_dai_mm, chieu_rong_khong_hong_mm) hoac None
# (nghia la ap dung chung, khong phan biet kich thuoc) -> dict cac thong so.
DAN_CAT_KHACH_DAC_BIET = {
    "nam bảo tín": {
        (850, 550): {
            "toc_do_may": "80 ± 10",
            "toc_do_keo": "50 ± 5",
            "lai_cuon": "630 ± 10",
            "luc_ep_duong_dan": "2 ± 1",
            "dieu_chinh_duong": "250 ± 10",  # theo xac nhan rieng ngay 2026-07-30
        },
    },
    "neo nam việt": {
        # TSVH DAN CAT muc "I. Neo Nam Viet" - bo sung them cap (dai,rong) khi gap don moi
    },
    "samyang": {
        (790, 550): {
            "toc_do_may": "70 ± 10", "toc_do_keo": "30 ± 5",
            "lai_cuon": "630 ± 10", "luc_ep_duong_dan": "2 ± 1",
            "dieu_chinh_duong": "250 ± 10",
        },
    },
}

# Bang TSVH CHUNG (khong phan biet khach hang) - trich tu sheet "TSDCA1" trong
# TSVH DAN CAT.xlsx, doc lai ngay 2026-07-30. Dung khi khong co muc rieng cho
# khach hang (hoac co khach nhung sai kich thuoc).
#
# Toc do may / Toc do keo: tra theo CHIEU DAI BAO (mm).
TOC_DO_MAY_GENERIC_RANGES = [
    (500, 700, "70 ± 10"),
    (710, 800, "80 ± 10"),
    (810, 900, "90 ± 10"),
    (910, 1100, "100 ± 10"),
]
TOC_DO_KEO_GENERIC_RANGES = [
    (500, 700, "27 ± 7"),
    (710, 800, "32 ± 7"),
    (810, 900, "35 ± 7"),
    (910, 1100, "35 ± 7"),
]

# Lai cuon / Dieu chinh duong: tra theo cap (NGANG mm, HONG mm) dung dinh dang
# "(ngang + hong)" trong TSVH goc. CHU Y: TSVH goc con 3 muc cuoi dung dinh dang
# "A x B" (vd "500 x 920") thay vi "A + B" - co the la quy uoc khac (kich thuoc
# tong the?), KHONG dua vao day de tranh nham lan; neu don co kich thuoc gan
# 500x920 / 550x890 / 600x1000 thi hoi lai nguoi dung truoc khi dung so nao.
LAI_CUON_GENERIC = {
    (300, 100): "408 ±10", (310, 90): "405 ±10", (320, 80): "420 ±10",
    (335, 80): "440 ±10", (340, 105): "490 ±10", (350, 100): "495 ±10",
    (370, 80): "505 ±10", (400, 100): "560 ±10", (420, 80): "560 ±10",
    (450, 100): "650 ±10", (450, 120): "700 ±10", (460, 100): "620 ±10",
    (480, 100): "635 ±10", (480, 120): "690 ±10", (450, 150): "715 ±10",
}
DIEU_CHINH_DUONG_GENERIC = {
    (320, 80): "90 ±10", (335, 80): "100 ±10", (340, 105): "120 ±10",
    (350, 100): "125 ±10", (370, 80): "135 ±10", (400, 100): "180 ±10",
    (420, 80): "190 ±10", (450, 120): "240 ±10", (460, 100): "250 ±10",
    (480, 100): "260 ±10", (480, 120): "260 ±10", (500, 100): "282 ±10",
    (450, 150): "240 ±10", (320, 100): "80 ±10", (450, 100): "240 ±10",
}
# Luc ep duong dan: theo loai bao ("KP, in offset: 2±1" / "OPP: 5±1" trong TSVH goc).
LUC_EP_DUONG_DAN_GENERIC = {
    "KP": "2 ± 1",
    "OFFSET": "2 ± 1",
    "OPP": "5 ± 1",
    "BOPP": "5 ± 1",   # bao_type dung trong code la "BOPP", khong phai "OPP"
}


def _range_lookup(mm, ranges):
    for lo, hi, val in ranges:
        if lo <= mm <= hi:
            return val
    return None


def dan_cat_tra_thong_so(customer_name, bao_type, chieu_dai_mm, ngang_mm, hong_mm):
    """
    Tra thong so van hanh Dan Cat. Thu tu uu tien:
      1. Muc khach hang dac biet (dung ten + dung kich thuoc) - neu co.
      2. Neu khong co (sai khach hoac sai kich thuoc) -> dung Bang TSVH CHUNG
         (TSDCA1), tra theo chieu dai bao / cap (ngang,hong) / loai bao.
    Tra ve (params_dict, source) voi source la "khach_dac_biet" | "chung".
    Truong nao khong tra duoc (kha ca o bang chung) se la None trong dict ->
    Claude de trong o do + to vang, KHONG tu noi suy them.
    """
    key = (int(round(chieu_dai_mm)), int(round(ngang_mm)))
    cust_key = unicodedata.normalize("NFC", (customer_name or "").strip().lower())
    for name, sizes in DAN_CAT_KHACH_DAC_BIET.items():
        if name in cust_key:
            if key in sizes:
                return dict(sizes[key]), "khach_dac_biet"
            break  # co khach nhung sai kich thuoc -> roi xuong dung bang chung

    dai = int(round(chieu_dai_mm))
    ng_hong_key = (int(round(ngang_mm)), int(round(hong_mm)))
    params = {
        "toc_do_may": _range_lookup(dai, TOC_DO_MAY_GENERIC_RANGES),
        "toc_do_keo": _range_lookup(dai, TOC_DO_KEO_GENERIC_RANGES),
        "lai_cuon": LAI_CUON_GENERIC.get(ng_hong_key),
        "luc_ep_duong_dan": LUC_EP_DUONG_DAN_GENERIC.get((bao_type or "").upper()),
        "dieu_chinh_duong": DIEU_CHINH_DUONG_GENERIC.get(ng_hong_key),
    }
    return params, "chung"


# Do sau toi da cua dia xep hong moi ben (Dan 2) - tra theo Hong (mm), trich tu
# TSVH DAN CAT.xlsx sheet TSDCA1, muc 15. Dung CHO CA BOPP VA KP (da doi chieu
# hai template, cung o dong 37/D37). Chi tra khop CHINH XAC gia tri Hong, KHONG
# noi suy - neu Hong khong khop muc nao trong bang, tra ve None (de trong, hoi
# lai) thay vi doan gan dung.
DO_SAU_DIA_XEP_HONG = {
    60: "30 ±3", 70: "35 ±3", 80: "40 ±3", 90: "45 ±3",
    100: "50 ±3", 110: "55 ±3", 120: "60 ±3", 150: "75 ±3",
}


def dan_cat_do_sau_dia_xep_hong(hong_mm):
    """Tra 'Do sau toi da cua dia xep hong moi ben' theo Hong (mm). None neu
    khong khop dung gia tri nao trong bang TSVH."""
    return DO_SAU_DIA_XEP_HONG.get(int(round(hong_mm)))


# ---------------------------------------------------------------------------
# THOI - tra theo kho tui long (chieu rong, mm), dung chung cho moi loai bao.
THOI_KHO_RANGES = [
    (400, 440, "50 ± 5", "30 ± 5"),
    (450, 490, "53 ± 5", "33 ± 5"),
    (500, 540, "55 ± 5", "35 ± 5"),
    (550, 590, "55 ± 5", "32 ± 5"),
    (600, 600, "60 ± 5", "28 ± 5"),
]


def thoi_tra_toc_do(kho_tui_long_mm):
    """Tra ve (toc_do_dun_keo, toc_do_keo_bong) theo kho tui long (mm)."""
    k = kho_tui_long_mm
    for lo, hi, dun_keo, keo_bong in THOI_KHO_RANGES:
        if lo <= k <= hi:
            return dun_keo, keo_bong
    return None, None


def do_day_tui_long_micron(khoi_luong_g, rong_cm, dai_cm):
    """Do day (micron) = KL(g)*10000/rong(cm)/dai(cm)/2/0.93, lam tron so nguyen."""
    return round(khoi_luong_g * 10000 / rong_cm / dai_cm / 2 / 0.93)
