# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Vài phép xử lý ảnh nhỏ bằng numpy + Pillow (thay OpenCV - bỏ được ~110 MB khỏi bản exe).
Ảnh là mảng numpy uint8 2 chiều (ảnh xám / ảnh nhị phân 0-255)."""
import numpy as np
from PIL import Image, ImageFilter


def otsu_threshold(a):
    """Ngưỡng Otsu của ảnh xám (giống cv2.THRESH_OTSU)."""
    hist = np.bincount(a.ravel(), minlength=256).astype(np.float64)
    total = hist.sum()
    if total == 0:
        return 0
    w0 = np.cumsum(hist)
    m = np.cumsum(hist * np.arange(256))
    w1 = total - w0
    with np.errstate(divide="ignore", invalid="ignore"):
        between = (m[-1] * w0 - m * total) ** 2 / (w0 * w1)
    between[~np.isfinite(between)] = 0
    return int(np.argmax(between))


def rotate(bw, deg):
    """Xoay quanh tâm, giữ kích thước, lân cận gần nhất, nền 0 (dương = ngược chiều kim đồng hồ)."""
    return np.asarray(Image.fromarray(bw).rotate(deg, resample=Image.NEAREST, fillcolor=0))


def _box_mean(img, k):
    """Trung bình cửa sổ k x k, viền lặp lại điểm ảnh biên (như cv2.boxFilter / BORDER_REPLICATE)."""
    r = k // 2
    p = np.pad(img.astype(np.float64), r, mode="edge")
    s = np.pad(p.cumsum(0).cumsum(1), ((1, 0), (1, 0)))
    H, W = img.shape
    tot = s[k:k + H, k:k + W] - s[:H, k:k + W] - s[k:k + H, :W] + s[:H, :W]
    return tot / (k * k)


def adaptive_mean_inv(img, block=25, c=15):
    """Như cv2.adaptiveThreshold(MEAN_C, THRESH_BINARY_INV): mực (tối hơn trung bình vùng - c) = 255."""
    t = _box_mean(img, block) - c
    return np.where(img.astype(np.float64) <= t, 255, 0).astype(np.uint8)


def _run_window(mask, length, axis, need_all):
    """Cửa sổ trượt 1 chiều dài `length` (tâm như OpenCV): need_all = co (erode), ngược lại = giãn (dilate)."""
    a = np.moveaxis(mask, axis, -1).astype(np.int32)
    before = length // 2
    after = length - 1 - before
    pad_val = 1 if need_all else 0                       # erode coi ngoài ảnh là "có", dilate là "không"
    p = np.pad(a, [(0, 0)] * (a.ndim - 1) + [(before, after)], constant_values=pad_val)
    cs = np.concatenate([np.zeros(a.shape[:-1] + (1,), np.int32), p.cumsum(-1)], -1)
    win = cs[..., length:] - cs[..., :-length]
    out = win == length if need_all else win > 0
    return np.moveaxis(out, -1, axis)


def open_line(binary, length, axis):
    """Mở hình thái học với đoạn thẳng dài `length` (axis=1 ngang, 0 dọc): chỉ giữ nét thẳng đủ dài."""
    m = binary > 0
    return _run_window(_run_window(m, length, axis, True), length, axis, False)


def dilate3(mask):
    """Giãn 3x3."""
    img = Image.fromarray(np.where(mask, 255, 0).astype(np.uint8))
    return np.asarray(img.filter(ImageFilter.MaxFilter(3))) > 0


def median3(a):
    """Lọc trung vị 3x3 (như cv2.medianBlur(a, 3))."""
    return np.asarray(Image.fromarray(a).filter(ImageFilter.MedianFilter(3)))


def label_components(mask, connectivity=8):
    """Gán nhãn vùng liên thông 8 (hoặc 4) hướng. Trả về (số nhãn kể cả nền 0, ảnh nhãn, diện tích từng nhãn).
    Làm theo các đoạn chạy (run) trên từng hàng + union-find -> đủ nhanh cho ảnh con dấu / trang scan."""
    diag = 1 if connectivity == 8 else 0                # 8 hướng: run chạm chéo cũng tính là dính
    H, W = mask.shape
    m = mask.astype(np.int8)
    d = np.diff(np.pad(m, ((0, 0), (1, 1))), axis=1)
    rows_s, cols_s = np.nonzero(d == 1)
    _, cols_e = np.nonzero(d == -1)                     # cột kết thúc (không gồm) - cùng thứ tự với điểm bắt đầu
    n = len(rows_s)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    row_start = np.searchsorted(rows_s, np.arange(H + 1))
    for y in range(1, H):
        a0, a1 = row_start[y - 1], row_start[y]          # các run của hàng trên
        b0, b1 = row_start[y], row_start[y + 1]          # các run của hàng này
        i = a0
        for j in range(b0, b1):
            s, e = cols_s[j], cols_e[j]
            while i < a1 and cols_e[i] + diag <= s:      # run trên kết thúc trước -> không dính
                i += 1
            k = i
            while k < a1 and cols_s[k] < e + diag:       # chồng lấn (8 hướng: kể cả chéo)
                ra, rb = find(k), find(j)
                if ra != rb:
                    parent[rb] = ra
                k += 1
    roots = np.array([find(x) for x in range(n)], np.int64) if n else np.zeros(0, np.int64)
    uniq, lab_of_run = np.unique(roots, return_inverse=True)
    labels = np.zeros((H, W), np.int32)
    lengths = (cols_e - cols_s).astype(np.int64)
    for r in range(n):
        labels[rows_s[r], cols_s[r]:cols_e[r]] = lab_of_run[r] + 1
    areas = np.zeros(len(uniq) + 1, np.int64)
    np.add.at(areas, lab_of_run + 1, lengths)
    return len(uniq) + 1, labels, areas
