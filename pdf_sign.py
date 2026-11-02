# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Ký số PDF bằng chứng thư trong kho Windows (USB Token VNPT-CA, Viettel-CA, BKAV, FPT...).

Token đăng ký chứng thư vào kho "Personal" của Windows. Khi ký, Windows gọi driver của token
(KSP/CSP) và driver tự hiện hộp thoại nhập PIN. Phần định dạng PDF/CMS do pyHanko đảm nhiệm.
"""
import ctypes
import hashlib
import io
import logging
import os
from ctypes import wintypes
from datetime import datetime, timezone

from asn1crypto import algos, x509

# Log của pyHanko / fontTools rất nhiều (vd "chứng thư tự ký") - chỉ giữ lỗi nghiêm trọng
for _name in ("pyhanko", "pyhanko_certvalidator", "fontTools"):
    logging.getLogger(_name).setLevel(logging.CRITICAL)

# ---------------------------------------------------------------------------
#  Windows CryptoAPI / CNG
# ---------------------------------------------------------------------------
crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
ncrypt = ctypes.WinDLL("ncrypt", use_last_error=True)


class CERT_CONTEXT(ctypes.Structure):
    _fields_ = [("dwCertEncodingType", wintypes.DWORD), ("pbCertEncoded", ctypes.POINTER(ctypes.c_ubyte)),
                ("cbCertEncoded", wintypes.DWORD), ("pCertInfo", ctypes.c_void_p), ("hCertStore", ctypes.c_void_p)]


PCCERT_CONTEXT = ctypes.POINTER(CERT_CONTEXT)


class BCRYPT_PKCS1_PADDING_INFO(ctypes.Structure):
    _fields_ = [("pszAlgId", ctypes.c_wchar_p)]


crypt32.CertOpenSystemStoreW.restype = ctypes.c_void_p
crypt32.CertOpenSystemStoreW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p]
crypt32.CertEnumCertificatesInStore.restype = PCCERT_CONTEXT
crypt32.CertEnumCertificatesInStore.argtypes = [ctypes.c_void_p, PCCERT_CONTEXT]
crypt32.CertDuplicateCertificateContext.restype = PCCERT_CONTEXT
crypt32.CertDuplicateCertificateContext.argtypes = [PCCERT_CONTEXT]
crypt32.CertFreeCertificateContext.argtypes = [PCCERT_CONTEXT]
crypt32.CertCloseStore.argtypes = [ctypes.c_void_p, wintypes.DWORD]
crypt32.CertGetCertificateContextProperty.argtypes = [PCCERT_CONTEXT, wintypes.DWORD, ctypes.c_void_p,
                                                      ctypes.POINTER(wintypes.DWORD)]
crypt32.CryptAcquireCertificatePrivateKey.argtypes = [PCCERT_CONTEXT, wintypes.DWORD, ctypes.c_void_p,
                                                      ctypes.POINTER(ctypes.c_size_t), ctypes.POINTER(wintypes.DWORD),
                                                      ctypes.POINTER(wintypes.BOOL)]
ncrypt.NCryptSignHash.argtypes = [ctypes.c_size_t, ctypes.c_void_p, ctypes.c_char_p, wintypes.DWORD,
                                  ctypes.c_char_p, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), wintypes.DWORD]
ncrypt.NCryptSignHash.restype = ctypes.c_long
ncrypt.NCryptFreeObject.argtypes = [ctypes.c_size_t]
advapi32.CryptCreateHash.argtypes = [ctypes.c_size_t, wintypes.DWORD, ctypes.c_size_t, wintypes.DWORD,
                                     ctypes.POINTER(ctypes.c_size_t)]
advapi32.CryptSetHashParam.argtypes = [ctypes.c_size_t, wintypes.DWORD, ctypes.c_char_p, wintypes.DWORD]
advapi32.CryptSignHashW.argtypes = [ctypes.c_size_t, wintypes.DWORD, ctypes.c_wchar_p, wintypes.DWORD,
                                    ctypes.c_char_p, ctypes.POINTER(wintypes.DWORD)]
advapi32.CryptDestroyHash.argtypes = [ctypes.c_size_t]
advapi32.CryptSetProvParam.argtypes = [ctypes.c_size_t, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
ncrypt.NCryptSetProperty.argtypes = [ctypes.c_size_t, ctypes.c_wchar_p, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD]
advapi32.CryptReleaseContext.argtypes = [ctypes.c_size_t, wintypes.DWORD]

CERT_KEY_PROV_INFO_PROP_ID = 2
CRYPT_ACQUIRE_CACHE_FLAG = 0x1
CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG = 0x00020000
CERT_NCRYPT_KEY_SPEC = 0xFFFFFFFF
BCRYPT_PAD_PKCS1 = 0x2
HP_HASHVAL = 0x2
CALG = {"sha1": 0x8004, "sha256": 0x800C, "sha384": 0x800D, "sha512": 0x800E}

ERRORS = {
    0x8010006E: "Đã huỷ nhập PIN.",
    0x80100069: "Không tìm thấy token — hãy cắm USB Token.",
    0x8010000C: "Không tìm thấy thẻ / token trong đầu đọc.",
    0x80090016: "Không tìm thấy khoá bí mật (token chưa cắm hoặc sai token).",
    0x8009000D: "Không tìm thấy khoá bí mật của chứng thư.",
    0x80090027: "Tham số ký không hợp lệ.",
    0x800704C7: "Đã huỷ thao tác.",
    0x80090029: "Token không hỗ trợ thuật toán ký SHA-256.",
    0x8009001D: "Driver token chưa cài đặt / không tải được.",
    0x80070005: "Bị từ chối truy cập khoá (PIN sai hoặc token bị khoá).",
}


PP_CLIENT_HWND = 1
PARENT_HWND = None          # cửa sổ cha cho hộp thoại nhập PIN của token (đặt từ giao diện)


class SignError(Exception):
    pass


def _err(code, where):
    code &= 0xFFFFFFFF
    return SignError(ERRORS.get(code, f"Lỗi ký số ({where}): 0x{code:08X}"))


class WinCert:
    """Chứng thư trong kho Windows có khoá bí mật."""

    def __init__(self, ctx):
        self.ctx = ctx
        der = ctypes.string_at(ctx.contents.pbCertEncoded, ctx.contents.cbCertEncoded)
        self.cert = x509.Certificate.load(der)
        self.der = der

    # --- thông tin hiển thị ---
    @staticmethod
    def _name(name, key="common_name"):
        try:
            return name.native.get(key) or ""
        except Exception:
            return ""

    @property
    def subject(self):
        return self._name(self.cert.subject) or self.cert.subject.human_friendly

    @property
    def organization(self):
        return self._name(self.cert.subject, "organization_name")

    @property
    def issuer(self):
        return self._name(self.cert.issuer) or self.cert.issuer.human_friendly

    @property
    def not_after(self):
        return self.cert["tbs_certificate"]["validity"]["not_after"].native

    @property
    def not_before(self):
        return self.cert["tbs_certificate"]["validity"]["not_before"].native

    @property
    def expired(self):
        now = datetime.now(timezone.utc)
        return not (self.not_before <= now <= self.not_after)

    @property
    def serial(self):
        return format(self.cert.serial_number, "X")

    @property
    def key_algo(self):
        return self.cert.public_key.algorithm      # 'rsa' | 'ec' | ...

    @property
    def key_bytes(self):
        try:
            return (self.cert.public_key.bit_size + 7) // 8
        except Exception:
            return 256

    @property
    def can_sign(self):
        try:
            ku = self.cert.key_usage_value
            if ku is None:
                return True
            return bool({"digital_signature", "non_repudiation"} & set(ku.native))
        except Exception:
            return True

    def label(self):
        org = f" — {self.organization}" if self.organization else ""
        return f"{self.subject}{org}  (CA: {self.issuer}, hết hạn {self.not_after:%d/%m/%Y})"

    # --- ký ---
    def sign_hash(self, digest, digest_algorithm="sha256"):
        key = ctypes.c_size_t()
        spec = wintypes.DWORD()
        free = wintypes.BOOL()
        hwnd = ctypes.c_void_p(PARENT_HWND) if PARENT_HWND else None
        if hwnd:
            advapi32.CryptSetProvParam(0, PP_CLIENT_HWND, ctypes.byref(hwnd), 0)
        if not crypt32.CryptAcquireCertificatePrivateKey(self.ctx, CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG, None,
                                                         ctypes.byref(key), ctypes.byref(spec), ctypes.byref(free)):
            code = ctypes.get_last_error()
            raise _err(code | 0x80070000 if code < 0x10000 else code, "lấy khoá")
        try:
            if spec.value == CERT_NCRYPT_KEY_SPEC:
                if hwnd:
                    ncrypt.NCryptSetProperty(key.value, "HWND Handle", ctypes.byref(hwnd), ctypes.sizeof(hwnd), 0)
                return self._sign_ncrypt(key.value, digest, digest_algorithm)
            return self._sign_csp(key.value, spec.value, digest, digest_algorithm)
        finally:
            if free.value:
                if spec.value == CERT_NCRYPT_KEY_SPEC:
                    ncrypt.NCryptFreeObject(key.value)
                else:
                    advapi32.CryptReleaseContext(key.value, 0)

    @staticmethod
    def _sign_ncrypt(key, digest, alg):
        pad = BCRYPT_PKCS1_PADDING_INFO(alg.upper())
        cb = wintypes.DWORD()
        st = ncrypt.NCryptSignHash(key, ctypes.byref(pad), digest, len(digest), None, 0, ctypes.byref(cb), BCRYPT_PAD_PKCS1)
        if st != 0:
            raise _err(st, "NCryptSignHash")
        buf = ctypes.create_string_buffer(cb.value)
        st = ncrypt.NCryptSignHash(key, ctypes.byref(pad), digest, len(digest), buf, cb, ctypes.byref(cb), BCRYPT_PAD_PKCS1)
        if st != 0:
            raise _err(st, "NCryptSignHash")
        return buf.raw[:cb.value]

    @staticmethod
    def _sign_csp(prov, spec, digest, alg):
        h = ctypes.c_size_t()
        if not advapi32.CryptCreateHash(prov, CALG[alg], 0, 0, ctypes.byref(h)):
            raise _err(ctypes.get_last_error() | 0x80090000, "CryptCreateHash")
        try:
            if not advapi32.CryptSetHashParam(h.value, HP_HASHVAL, digest, 0):
                raise _err(ctypes.get_last_error() | 0x80090000, "CryptSetHashParam")
            cb = wintypes.DWORD()
            if not advapi32.CryptSignHashW(h.value, spec, None, 0, None, ctypes.byref(cb)):
                raise _err(ctypes.get_last_error() | 0x80090000, "CryptSignHash")
            buf = ctypes.create_string_buffer(cb.value)
            if not advapi32.CryptSignHashW(h.value, spec, None, 0, buf, ctypes.byref(cb)):
                raise _err(ctypes.get_last_error() | 0x80090000, "CryptSignHash")
            return buf.raw[:cb.value][::-1]          # CryptoAPI trả về little-endian
        finally:
            advapi32.CryptDestroyHash(h.value)


def _store_certs(store_name):
    store = crypt32.CertOpenSystemStoreW(None, store_name)
    if not store:
        return []
    out = []
    ctx = crypt32.CertEnumCertificatesInStore(store, None)
    while ctx:
        out.append(crypt32.CertDuplicateCertificateContext(ctx))
        ctx = crypt32.CertEnumCertificatesInStore(store, ctx)
    crypt32.CertCloseStore(store, 0)
    return out


_SYSTEM_CERT_MARKERS = ("ms-organization", "login.windows.net", "microsoft intune", "windows hello",
                        "localhost", "adfs", "msi certificate")


def _is_system_cert(c):
    """Chứng thư của Windows / Entra ID / Intune... (không phải chứng thư ký văn bản)."""
    text = f"{c.cert.subject.human_friendly} {c.cert.issuer.human_friendly}".lower()
    return any(m in text for m in _SYSTEM_CERT_MARKERS)


def list_signing_certs(include_expired=False, include_self_signed=False):
    """Chứng thư có khoá bí mật trong kho Personal (gồm cả chứng thư từ USB Token).

    Bỏ qua chứng thư tự ký và chứng thư hệ thống - chứng thư số thật luôn do một CA cấp."""
    out = []
    for ctx in _store_certs("MY"):
        cb = wintypes.DWORD()
        if not crypt32.CertGetCertificateContextProperty(ctx, CERT_KEY_PROV_INFO_PROP_ID, None, ctypes.byref(cb)):
            continue
        try:
            c = WinCert(ctx)
        except Exception:
            continue
        if not c.can_sign or (c.expired and not include_expired) or _is_system_cert(c):
            continue
        if not include_self_signed and c.cert.self_signed != "no":
            continue
        out.append(c)
    out.sort(key=lambda c: c.not_after, reverse=True)
    return out


def build_chain(cert):
    """Tìm các chứng thư CA trung gian / gốc trong kho Windows để nhúng vào chữ ký."""
    pool = []
    for name in ("CA", "Root"):
        for ctx in _store_certs(name):
            try:
                pool.append(x509.Certificate.load(ctypes.string_at(ctx.contents.pbCertEncoded, ctx.contents.cbCertEncoded)))
            except Exception:
                pass
    chain, cur = [], cert
    for _ in range(6):
        if cur.self_signed != "no":
            break
        issuer = next((c for c in pool if c.subject == cur.issuer and c not in chain), None)
        if issuer is None:
            break
        chain.append(issuer)
        cur = issuer
    return chain


# ---------------------------------------------------------------------------
#  pyHanko
# ---------------------------------------------------------------------------
def _make_signer(wc):
    from pyhanko.sign import signers
    from pyhanko_certvalidator.registry import SimpleCertificateStore

    if wc.key_algo != "rsa":
        raise SignError("Chứng thư dùng khoá ECC — phiên bản này chỉ hỗ trợ khoá RSA (đa số token tại VN).")

    class WinStoreSigner(signers.Signer):
        def __init__(self):
            super().__init__(signing_cert=wc.cert,
                             cert_registry=SimpleCertificateStore.from_certs(build_chain(wc.cert)),
                             signature_mechanism=algos.SignedDigestAlgorithm({"algorithm": "sha256_rsa"}))

        async def async_sign_raw(self, data, digest_algorithm, dry_run=False):
            if dry_run:
                return bytes(wc.key_bytes)
            return wc.sign_hash(hashlib.new(digest_algorithm, data).digest(), digest_algorithm)

    return WinStoreSigner()


FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")


def sign_pdf(pdf_bytes, wc, page_index=0, box=None, reason="", location="", contact="", image_path=None,
             show_text=True, tsa_url=None, field_name=None, check_only=False):
    """Ký PDF. `box` = (x1, y1, x2, y2) theo toạ độ PDF (gốc dưới-trái), None = ký ẩn. Trả về bytes."""
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign import fields, signers, timestamps
    from pyhanko.stamp import TextStampStyle
    from pyhanko.pdf_utils.layout import SimpleBoxLayoutRule, AxisAlignment, Margins, InnerScaling

    w = IncrementalPdfFileWriter(io.BytesIO(pdf_bytes), strict=False)
    existing = set()
    try:
        existing = {f[0] for f in fields.enumerate_sig_fields(w.prev)}
    except Exception:
        pass
    k = 1
    while (field_name or f"ChuKySo{k}") in existing:
        k += 1
    field_name = field_name or f"ChuKySo{k}"

    meta = signers.PdfSignatureMetadata(field_name=field_name, reason=reason or None, location=location or None,
                                        contact_info=contact or None, md_algorithm="sha256", subfilter=fields.SigSeedSubFilter.PADES)
    if box:
        box = tuple(int(round(v)) for v in (min(box[0], box[2]), min(box[1], box[3]), max(box[0], box[2]), max(box[1], box[3])))
    spec = fields.SigFieldSpec(sig_field_name=field_name, on_page=page_index, box=box) if box else \
        fields.SigFieldSpec(sig_field_name=field_name)

    style = None
    tmp_pdf = None
    if box:
        # Hình ô chữ ký tự vẽ bằng PyMuPDF (font Arial tiếng Việt, vector), nhúng như một trang PDF nhỏ.
        # (Bộ dựng font của pyHanko sinh sai độ rộng ký tự với fontTools mới -> chữ bị giãn.)
        from pyhanko.pdf_utils.content import ImportedPdfPage
        import tempfile
        fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        lines = stamp_lines(wc, reason, location) if show_text and not check_only else []
        make_appearance(tmp_pdf, box[2] - box[0], box[3] - box[1], lines, None if check_only else image_path,
                        check_only)
        style = TextStampStyle(border_width=0, stamp_text="", background=ImportedPdfPage(tmp_pdf),
                               background_opacity=1,
                               background_layout=SimpleBoxLayoutRule(
                                   x_align=AxisAlignment.ALIGN_MID, y_align=AxisAlignment.ALIGN_MID,
                                   margins=Margins(0, 0, 0, 0), inner_content_scaling=InnerScaling.STRETCH_FILL))

    tsa = timestamps.HTTPTimeStamper(tsa_url) if tsa_url else None
    pdf_signer = signers.PdfSigner(meta, signer=_make_signer(wc), timestamper=tsa, stamp_style=style, new_field_spec=spec)
    out = io.BytesIO()
    try:
        pdf_signer.sign_pdf(w, output=out)
    finally:
        if tmp_pdf:
            try:
                os.remove(tmp_pdf)
            except OSError:
                pass
    return out.getvalue()


def stamp_lines(wc, reason="", location="", when=None):
    """Các dòng chữ trong ô chữ ký: [(nhãn, giá trị)]."""
    lines = [("Ký số bởi: ", wc.subject)]
    if wc.organization:
        lines.append(("Đơn vị: ", wc.organization))
    lines.append(("Ngày ký: ", (when or datetime.now()).strftime("%d/%m/%Y %H:%M:%S")))
    if reason:
        lines.append(("Lý do: ", reason))
    if location:
        lines.append(("Nơi ký: ", location))
    return lines


def make_appearance(path, w, h, lines, image_path=None, check_only=False):
    """Vẽ hình ô chữ ký (w x h điểm) ra file PDF 1 trang."""
    doc = appearance_doc(w, h, lines, image_path, check_only)
    doc.subset_fonts()
    doc.save(path, garbage=3, deflate=True)
    doc.close()


def appearance_doc(w, h, lines, image_path=None, check_only=False):
    """Tài liệu 1 trang chứa hình ô chữ ký (dùng để nhúng và để xem trước).
    check_only: chỉ vẽ dấu ✔ xanh (xác nhận đã ký) giữa ô, không khung, không chữ."""
    import html
    import pymupdf
    doc = pymupdf.open()
    page = doc.new_page(width=w, height=h)
    if check_only:
        s = min(w, h) * 0.86                              # dấu tick vector, nét tròn, cân giữa ô
        x0, y0 = (w - s) / 2, (h - s) / 2
        pts = [pymupdf.Point(x0 + s * 0.14, y0 + s * 0.54), pymupdf.Point(x0 + s * 0.40, y0 + s * 0.80),
               pymupdf.Point(x0 + s * 0.88, y0 + s * 0.20)]
        sh = page.new_shape()
        sh.draw_polyline(pts)
        sh.finish(color=(0.09, 0.64, 0.29), width=max(1.5, s * 0.13), lineCap=1, lineJoin=1, closePath=False)
        sh.commit()
        return doc
    page.draw_rect(pymupdf.Rect(0.5, 0.5, w - 0.5, h - 0.5), color=(0.13, 0.45, 0.25), width=1)
    x_text = 4
    if image_path:
        img_w = w * 0.36 if lines else w - 8
        page.insert_image(pymupdf.Rect(4, 4, 4 + img_w, h - 4), filename=image_path, keep_proportion=True)
        x_text = 8 + img_w
    if lines:
        body = "".join(f"<div><b>{html.escape(k)}</b>{html.escape(v)}</div>" for k, v in lines)
        # cỡ chữ: mỗi dòng nằm gọn trên một hàng nếu được (6 - 10pt), quá cao thì insert_htmlbox tự thu nhỏ
        reg = pymupdf.Font(fontfile=os.path.join(FONT_DIR, "arial.ttf"))
        bold = pymupdf.Font(fontfile=os.path.join(FONT_DIR, "arialbd.ttf"))
        widest = max(bold.text_length(k, 1) + reg.text_length(v, 1) for k, v in lines)
        fsize = max(6.0, min(10.0, (w - 3 - x_text - 2) / (widest * 1.1), (h - 6) / (len(lines) * 1.2)))
        css = ("@font-face {font-family: VN; src: url(arial.ttf);} "
               "@font-face {font-family: VN; src: url(arialbd.ttf); font-weight: bold;} "
               "* {font-family: VN; font-size: %.1fpt; color: #14532d; line-height: 1.2;}" % fsize)
        page.insert_htmlbox(pymupdf.Rect(x_text, 3, w - 3, h - 3), body, css=css,
                            archive=pymupdf.Archive(FONT_DIR), scale_low=0)
    return doc


def verify_pdf(pdf_bytes):
    """Kiểm tra toàn vẹn các chữ ký trong file. Trả về danh sách dict."""
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.sign.validation import validate_pdf_signature
    from pyhanko_certvalidator import ValidationContext

    r = PdfFileReader(io.BytesIO(pdf_bytes), strict=False)
    out = []
    for sig in r.embedded_signatures:
        info = dict(field=sig.field_name, signer="", time=None, intact=False, valid=False, covers_all=False, error="")
        try:
            if sig.sig_object.get("/Type") == "/DocTimeStamp":       # dấu thời gian tài liệu (RFC 3161)
                from pyhanko.sign.validation import validate_pdf_timestamp
                st = validate_pdf_timestamp(sig, validation_context=ValidationContext(allow_fetching=False))
                cert = st.signing_cert
                info.update(signer="Dấu thời gian: " + cert.subject.native.get("common_name", ""),
                            time=st.timestamp, intact=st.intact, valid=st.valid,
                            covers_all=st.coverage.name == "ENTIRE_FILE" if hasattr(st, "coverage") else False)
                out.append(info)
                continue
            st = validate_pdf_signature(sig, ValidationContext(allow_fetching=False))
            cert = st.signing_cert
            info.update(signer=cert.subject.native.get("common_name", cert.subject.human_friendly),
                        time=st.signer_reported_dt, intact=st.intact, valid=st.valid,
                        covers_all=st.coverage.name == "ENTIRE_FILE" if hasattr(st, "coverage") else False)
        except Exception as e:
            info["error"] = str(e)
        out.append(info)
    return out
