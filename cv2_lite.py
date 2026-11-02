# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Bản thay thế rất nhỏ cho OpenCV, chỉ gồm các hàm thư viện pdf2docx (PDF -> Word) dùng tới.
Bản exe không kèm OpenCV (~110 MB); install() gắn module này làm 'cv2' khi máy không có OpenCV thật.
Ảnh màu theo quy ước OpenCV: mảng numpy H x W x 3, thứ tự kênh BGR."""
import io
import sys

import numpy as np
from PIL import Image

import imgops

IMREAD_COLOR = 1
IMREAD_GRAYSCALE = 0
COLOR_BGR2GRAY = 6
THRESH_BINARY = 0
THRESH_BINARY_INV = 1
RETR_TREE = 3
CHAIN_APPROX_SIMPLE = 2
INTER_LINEAR = 1
INTER_NEAREST = 0


def install():
    """Dùng module này làm 'cv2' nếu không có OpenCV thật."""
    import importlib.util
    if "cv2" in sys.modules:
        return
    try:
        spec = importlib.util.find_spec("cv2")
    except (ImportError, ValueError):
        spec = None
    if spec is None:
        sys.modules["cv2"] = sys.modules[__name__]


def imdecode(buf, flags=IMREAD_COLOR):
    try:
        im = Image.open(io.BytesIO(np.asarray(buf, np.uint8).tobytes()))
    except Exception:
        return None
    if flags == IMREAD_GRAYSCALE:
        return np.array(im.convert("L"))
    return np.array(im.convert("RGB"))[:, :, ::-1].copy()


def imencode(ext, img, params=None):
    a = np.asarray(img)
    if a.ndim == 3 and a.shape[2] == 3:
        a = a[:, :, ::-1]
    elif a.ndim == 3 and a.shape[2] == 4:
        a = a[:, :, [2, 1, 0, 3]]
    out = io.BytesIO()
    Image.fromarray(np.ascontiguousarray(a)).save(out, "PNG" if ext.lower().endswith("png") else ext.strip(".").upper())
    return True, np.frombuffer(out.getvalue(), np.uint8)


def cvtColor(img, code):
    if code != COLOR_BGR2GRAY:
        raise NotImplementedError(f"cvtColor code {code}")
    f = np.asarray(img, np.float32)
    return np.clip(np.rint(f[..., 0] * 0.114 + f[..., 1] * 0.587 + f[..., 2] * 0.299), 0, 255).astype(np.uint8)


def threshold(src, thresh, maxval, type_):
    a = np.asarray(src)
    if type_ == THRESH_BINARY:
        dst = np.where(a > thresh, maxval, 0)
    elif type_ == THRESH_BINARY_INV:
        dst = np.where(a > thresh, 0, maxval)
    else:
        raise NotImplementedError(f"threshold type {type_}")
    return float(thresh), dst.astype(a.dtype)


def getRotationMatrix2D(center, angle, scale):
    a = np.deg2rad(angle)
    al, be = scale * np.cos(a), scale * np.sin(a)
    cx, cy = center
    return np.array([[al, be, (1 - al) * cx - be * cy],
                     [-be, al, be * cx + (1 - al) * cy]], np.float64)


def warpAffine(img, M, dsize, flags=INTER_LINEAR, borderValue=0):
    """Biến đổi affine (M: điểm nguồn -> điểm đích, như OpenCV); nền ngoài ảnh = borderValue."""
    a = np.asarray(img)
    W, H = int(dsize[0]), int(dsize[1])
    full = np.vstack([np.asarray(M, np.float64), [0, 0, 1]])
    inv = np.linalg.inv(full)[:2].ravel()                # PIL cần ánh xạ ngược: điểm đích -> điểm nguồn
    # OpenCV lấy tâm điểm ảnh tại toạ độ nguyên, PIL tại +0.5 -> bù nửa điểm ảnh
    a_, b_, c_, d_, e_, f_ = inv
    c_ = c_ - 0.5 * (a_ + b_) + 0.5
    f_ = f_ - 0.5 * (d_ + e_) + 0.5
    rgb = a.ndim == 3
    im = Image.fromarray(np.ascontiguousarray(a[:, :, ::-1] if rgb else a))
    fill = borderValue if not isinstance(borderValue, (tuple, list)) else tuple(borderValue[:3][::-1])
    if rgb and not isinstance(fill, tuple):
        fill = (fill,) * 3
    res = im.transform((W, H), Image.AFFINE, (a_, b_, c_, d_, e_, f_),
                       resample=Image.NEAREST if flags == INTER_NEAREST else Image.BILINEAR, fillcolor=fill)
    out = np.array(res)
    return out[:, :, ::-1].copy() if rgb else out


def findContours(image, mode, method):
    """Chỉ hỗ trợ RETR_TREE và chỉ để lấy khung bao + quan hệ cha con (pdf2docx chỉ cần vậy).
    Mỗi 'contour' trả về là khung (x, y, w, h) - dùng với boundingRect. hierarchy[0, i, 3] = chỉ số cha."""
    fg = np.asarray(image) > 0
    H, W = fg.shape
    if not fg.any():
        return (), None
    n_f, lab_f, _ = imgops.label_components(fg, 8)          # vật thể: 8 hướng (như OpenCV)
    n_b, lab_b, _ = imgops.label_components(~fg, 4)         # nền / lỗ: 4 hướng
    # vùng nền chạm mép ảnh = nền ngoài; còn lại là lỗ bên trong vật thể
    edge = np.unique(np.concatenate([lab_b[0], lab_b[-1], lab_b[:, 0], lab_b[:, -1]]))
    outer = set(int(x) for x in edge if x)

    def boxes(lab, n):
        b = np.full((n, 4), -1, np.int64)
        ys, xs = np.nonzero(lab)
        ls = lab[ys, xs]
        b[:, 0] = W
        b[:, 1] = H
        np.minimum.at(b[:, 0], ls, xs)
        np.minimum.at(b[:, 1], ls, ys)
        np.maximum.at(b[:, 2], ls, xs)
        np.maximum.at(b[:, 3], ls, ys)
        return b

    def first_pixels(lab, n):
        flat = lab.ravel()
        idx = np.full(n, -1, np.int64)
        nz = np.nonzero(flat)[0]
        # điểm đầu tiên theo thứ tự quét của mỗi nhãn
        rev = nz[::-1]
        idx[flat[rev]] = rev
        return idx

    bf, bb = boxes(lab_f, n_f), boxes(lab_b, n_b)
    pf, pb = first_pixels(lab_f, n_f), first_pixels(lab_b, n_b)
    contours, parents = [], []
    obj_index, hole_index = {}, {}
    holes = [h for h in range(1, n_b) if h not in outer]
    # cha của lỗ = vật thể ngay phía trên điểm đầu tiên của lỗ; cha của vật thể = vùng nền bên trái điểm đầu
    hole_owner = {}
    for h in holes:
        y, x = divmod(int(pb[h]), W)
        hole_owner[h] = int(lab_f[y - 1, x]) if y > 0 else 0
    obj_parent_bg = {}
    for o in range(1, n_f):
        y, x = divmod(int(pf[o]), W)
        obj_parent_bg[o] = int(lab_b[y, x - 1]) if x > 0 else 0
    for o in range(1, n_f):
        obj_index[o] = len(contours)
        x0, y0, x1, y1 = bf[o]
        contours.append((int(x0), int(y0), int(x1 - x0 + 1), int(y1 - y0 + 1)))
        parents.append(None)
    for h in holes:
        hole_index[h] = len(contours)
        x0, y0, x1, y1 = bb[h]                               # viền lỗ nằm trên điểm ảnh vật thể quanh lỗ
        x0, y0, x1, y1 = max(0, x0 - 1), max(0, y0 - 1), min(W - 1, x1 + 1), min(H - 1, y1 + 1)
        contours.append((int(x0), int(y0), int(x1 - x0 + 1), int(y1 - y0 + 1)))
        parents.append(None)
    for o, i in obj_index.items():
        bg = obj_parent_bg[o]
        parents[i] = hole_index.get(bg, -1)
    for h, i in hole_index.items():
        parents[i] = obj_index.get(hole_owner[h], -1)
    hier = np.full((1, len(contours), 4), -1, np.int32)
    hier[0, :, 3] = parents
    return tuple(contours), hier


def boundingRect(c):
    if isinstance(c, tuple) and len(c) == 4:
        return c
    pts = np.asarray(c).reshape(-1, 2)
    x0, y0 = pts.min(0)
    x1, y1 = pts.max(0)
    return int(x0), int(y0), int(x1 - x0 + 1), int(y1 - y0 + 1)


def rectangle(img, pt1, pt2, color, thickness=1):
    return img                                               # chỉ dùng khi gỡ lỗi trong pdf2docx


def imshow(*a, **k):
    pass


def waitKey(*a, **k):
    return -1
