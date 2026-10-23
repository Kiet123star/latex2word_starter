"""`backend/pdf_geometry.py` — MỘT quy ước toạ độ cho toàn bộ nhánh PDF.

Hệ chuẩn nội bộ: đơn vị point (1/72 in), gốc TRÊN-TRÁI của trang HIỂN THỊ (đã áp CropBox và /Rotate),
trục y tăng XUỐNG.  Đây cũng chính là hệ pixel khi render bằng pypdfium2 (×dpi/72), nên nhận dạng và
cắt vùng dùng chung một phép biến đổi.

* `user_to_display(x, y)`  : toạ độ người dùng PDF (y lên, gốc MediaBox) → hệ chuẩn.
* `plumber_to_display(x, top)`: toạ độ pdfplumber/pdfminer (đã xoay theo /Rotate, gốc trên-trái MediaBox)
  → hệ chuẩn.  pdfplumber báo `page.cropbox` sai khi trang xoay 180/270 nên không dùng giá trị đó.
* `display_rect_to_user(rect)`: ngược lại — để tạo PDF một vùng (MediaBox = vùng) khi cắt ảnh vector.

Ma trận CTM của nội dung và Form XObject lồng nhau đã được pdfminer/pdfium áp trước khi tới đây.
"""
from __future__ import annotations

from dataclasses import dataclass

__all__ = ["PageGeom"]

Rect = tuple[float, float, float, float]


@dataclass
class PageGeom:
    mediabox: Rect          # (x0, y0, x1, y1) user space
    cropbox: Rect           # đã giao với mediabox
    rotation: int           # 0 | 90 | 180 | 270 (theo chiều kim đồng hồ khi hiển thị)

    @classmethod
    def create(cls, mediabox, cropbox, rotation) -> "PageGeom":
        mb = tuple(float(v) for v in mediabox)
        mb = (min(mb[0], mb[2]), min(mb[1], mb[3]), max(mb[0], mb[2]), max(mb[1], mb[3]))
        cb = tuple(float(v) for v in (cropbox or mediabox))
        cb = (min(cb[0], cb[2]), min(cb[1], cb[3]), max(cb[0], cb[2]), max(cb[1], cb[3]))
        cb = (max(cb[0], mb[0]), max(cb[1], mb[1]), min(cb[2], mb[2]), min(cb[3], mb[3]))
        if cb[2] <= cb[0] or cb[3] <= cb[1]:
            cb = mb
        rot = int(rotation or 0) % 360
        rot = rot - rot % 90
        return cls(mb, cb, rot)

    @property
    def width(self) -> float:
        w, h = self.cropbox[2] - self.cropbox[0], self.cropbox[3] - self.cropbox[1]
        return h if self.rotation in (90, 270) else w

    @property
    def height(self) -> float:
        w, h = self.cropbox[2] - self.cropbox[0], self.cropbox[3] - self.cropbox[1]
        return w if self.rotation in (90, 270) else h

    # ── user space → hệ chuẩn
    def user_to_display(self, x: float, y: float) -> tuple[float, float]:
        cx0, cy0, cx1, cy1 = self.cropbox
        r = self.rotation
        if r == 0:
            return x - cx0, cy1 - y
        if r == 90:
            return y - cy0, x - cx0
        if r == 180:
            return cx1 - x, y - cy0
        return cy1 - y, cx1 - x  # 270

    def display_to_user(self, X: float, Y: float) -> tuple[float, float]:
        cx0, cy0, cx1, cy1 = self.cropbox
        r = self.rotation
        if r == 0:
            return cx0 + X, cy1 - Y
        if r == 90:
            return cx0 + Y, cy0 + X
        if r == 180:
            return cx1 - X, cy0 + Y
        return cx1 - Y, cy1 - X

    # ── pdfplumber (pdfminer đã xoay, gốc trên-trái của MediaBox đã xoay) → user space
    def plumber_to_user(self, xp: float, tp: float) -> tuple[float, float]:
        x0, y0, x1, y1 = self.mediabox
        r = self.rotation
        if r == 0:
            return xp + x0, y1 - tp
        if r == 90:
            return tp + x0, xp + y0
        if r == 180:
            return x1 - xp, tp + y0
        return x1 - tp, y1 - xp  # 270

    def plumber_to_display(self, xp: float, tp: float) -> tuple[float, float]:
        return self.user_to_display(*self.plumber_to_user(xp, tp))

    def plumber_rect(self, x0: float, top: float, x1: float, bottom: float) -> Rect:
        a = self.plumber_to_display(x0, top)
        b = self.plumber_to_display(x1, bottom)
        return (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))

    def user_rect(self, l: float, b: float, r: float, t: float) -> Rect:
        p = self.user_to_display(l, b)
        q = self.user_to_display(r, t)
        return (min(p[0], q[0]), min(p[1], q[1]), max(p[0], q[0]), max(p[1], q[1]))

    def display_rect_to_user(self, rect: Rect) -> Rect:
        a = self.display_to_user(rect[0], rect[1])
        b = self.display_to_user(rect[2], rect[3])
        return (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))

    def clip(self, rect: Rect) -> Rect:
        return (max(0.0, rect[0]), max(0.0, rect[1]), min(self.width, rect[2]), min(self.height, rect[3]))

    def to_dict(self) -> dict:
        return {"mediabox": list(self.mediabox), "cropbox": list(self.cropbox), "rotation": self.rotation,
                "width": round(self.width, 2), "height": round(self.height, 2)}
