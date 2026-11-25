"""Dựng Native PowerPoint (.pptx) 14 slides TechRecruit với thẻ Bento chỉnh sửa được,
ảnh Retina thật từ presentation/assets/ và hiệu ứng Animation bay vào tuần tự.
"""
import importlib
import subprocess
import sys
from pathlib import Path


def ensure_pptx() -> None:
    """Cài đặt python-pptx nếu chưa có."""
    try:
        importlib.import_module("pptx")
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx"])


def main() -> None:
    ensure_pptx()

    from lxml import etree
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt

    base_dir = Path(__file__).resolve().parent
    assets_dir = base_dir / "assets"
    output_pptx = Path.home() / "Documents" / "TechRecruit_Presentation.pptx"
    output_pptx.parent.mkdir(parents=True, exist_ok=True)

    # Bảng màu đồng bộ TechRecruit
    BG_DARK = RGBColor(0x06, 0x0D, 0x0C)
    CARD_BG = RGBColor(0x0C, 0x1A, 0x18)
    CARD_HIGHLIGHT = RGBColor(0x0F, 0x2B, 0x26)
    BORDER_EMERALD = RGBColor(0x10, 0xB9, 0x81)
    BORDER_SUBTLE = RGBColor(0x1E, 0x4A, 0x42)
    MINT = RGBColor(0x34, 0xD3, 0x99)
    CYAN = RGBColor(0x22, 0xD3, 0xEE)
    WHITE = RGBColor(0xF8, 0xFA, 0xFC)
    SLATE_LIGHT = RGBColor(0xCB, 0xD5, 0xE1)
    SLATE_MUTED = RGBColor(0x94, 0xA3, 0xB8)
    CHROME_BG = RGBColor(0x13, 0x24, 0x22)
    DOT_RED = RGBColor(0xFF, 0x5F, 0x56)
    DOT_YELLOW = RGBColor(0xFF, 0xBD, 0x2E)
    DOT_GREEN = RGBColor(0x27, 0xC9, 0x3F)

    FONT_HEAD = "Segoe UI"
    FONT_BODY = "Segoe UI"
    FONT_MONO = "Consolas"

    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def set_dark_bg(slide) -> None:
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = BG_DARK

        # Dải nhấn emerald mảnh trên cùng
        top_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.05)
        )
        top_line.fill.solid()
        top_line.fill.fore_color.rgb = BORDER_EMERALD
        top_line.line.fill.background()

    def add_header(slide, pill_text: str, title_text: str, slide_num: int, total: int = 14) -> list:
        anim_ids = []

        # Module pill
        pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(0.38), Inches(3.9), Inches(0.36)
        )
        pill.adjustments[0] = 0.5
        pill.fill.solid()
        pill.fill.fore_color.rgb = RGBColor(0x0D, 0x29, 0x24)
        pill.line.color.rgb = BORDER_EMERALD
        pill.line.width = Pt(1)
        tf = pill.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.14)
        tf.margin_top = Inches(0.04)
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = pill_text
        r.font.name = FONT_HEAD
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = MINT
        anim_ids.append(pill.shape_id)

        # Title
        tb = slide.shapes.add_textbox(Inches(0.65), Inches(0.82), Inches(12.0), Inches(0.7))
        tf_t = tb.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = 0
        tf_t.margin_top = 0
        p_t = tf_t.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title_text
        r_t.font.name = FONT_HEAD
        r_t.font.size = Pt(28)
        r_t.font.bold = True
        r_t.font.color.rgb = WHITE
        anim_ids.append(tb.shape_id)

        # Footer (không cần animate để luôn cố định)
        ft_l = slide.shapes.add_textbox(Inches(0.65), Inches(6.98), Inches(4.0), Inches(0.32))
        p_fl = ft_l.text_frame.paragraphs[0]
        r_fl = p_fl.add_run()
        r_fl.text = "TechRecruit · Nhóm 5 · HN-KS25-CNTT5"
        r_fl.font.name = FONT_BODY
        r_fl.font.size = Pt(11)
        r_fl.font.color.rgb = SLATE_MUTED

        ft_r = slide.shapes.add_textbox(Inches(11.3), Inches(6.98), Inches(1.4), Inches(0.32))
        p_fr = ft_r.text_frame.paragraphs[0]
        p_fr.alignment = PP_ALIGN.RIGHT
        r_fr = p_fr.add_run()
        r_fr.text = f"{slide_num:02d} / {total:02d}"
        r_fr.font.name = FONT_MONO
        r_fr.font.size = Pt(11.5)
        r_fr.font.bold = True
        r_fr.font.color.rgb = MINT

        return anim_ids

    def add_bento_card(
        slide,
        left: float,
        top: float,
        width: float,
        height: float,
        title: str,
        body_lines: list[str],
        tag: str = "",
        metric: str = "",
        highlight: bool = False,
    ) -> int:
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        card.adjustments[0] = 0.06
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_HIGHLIGHT if highlight else CARD_BG
        card.line.color.rgb = BORDER_EMERALD if highlight else BORDER_SUBTLE
        card.line.width = Pt(1.4 if highlight else 1.0)

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.24)
        tf.margin_right = Inches(0.22)
        tf.margin_top = Inches(0.2)
        tf.margin_bottom = Inches(0.18)

        first = True
        if metric:
            p_m = tf.paragraphs[0]
            first = False
            p_m.space_after = Pt(4)
            rm = p_m.add_run()
            rm.text = metric
            rm.font.name = FONT_HEAD
            rm.font.size = Pt(28)
            rm.font.bold = True
            rm.font.color.rgb = MINT

        if title:
            p_h = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p_h.space_after = Pt(6)
            rh = p_h.add_run()
            rh.text = title
            rh.font.name = FONT_HEAD
            rh.font.size = Pt(15.5)
            rh.font.bold = True
            rh.font.color.rgb = WHITE

        for line in body_lines:
            p_b = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p_b.space_after = Pt(5)
            rb = p_b.add_run()
            rb.text = f"• {line}" if len(body_lines) > 1 else line
            rb.font.name = FONT_BODY
            rb.font.size = Pt(12.5)
            rb.font.color.rgb = SLATE_LIGHT

        if tag:
            p_tag = tf.add_paragraph()
            p_tag.space_before = Pt(8)
            rt = p_tag.add_run()
            rt.text = f"[ {tag} ]"
            rt.font.name = FONT_MONO
            rt.font.size = Pt(10.5)
            rt.font.bold = True
            rt.font.color.rgb = CYAN

        return card.shape_id

    def add_browser_mockup(
        slide,
        left: float,
        top: float,
        width: float,
        img_name: str,
        url_text: str,
        badge_text: str = "",
        aspect_16_10: bool = False,
    ) -> list[int]:
        ids = []
        bar_h = 0.34
        img_h = width * (10 / 16 if aspect_16_10 else 9 / 16)

        # Khung viền ngoài
        frame = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left),
            Inches(top),
            Inches(width),
            Inches(bar_h + img_h + 0.04),
        )
        frame.adjustments[0] = 0.03
        frame.fill.solid()
        frame.fill.fore_color.rgb = CHROME_BG
        frame.line.color.rgb = BORDER_EMERALD
        frame.line.width = Pt(1.3)
        ids.append(frame.shape_id)

        # 3 nút tròn macOS
        for idx, color in enumerate((DOT_RED, DOT_YELLOW, DOT_GREEN)):
            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(left + 0.14 + idx * 0.16),
                Inches(top + 0.11),
                Inches(0.11),
                Inches(0.11),
            )
            dot.fill.solid()
            dot.fill.fore_color.rgb = color
            dot.line.fill.background()

        # URL pill
        url_box = slide.shapes.add_textbox(
            Inches(left + 0.7), Inches(top + 0.05), Inches(width - 0.85), Inches(0.24)
        )
        p_u = url_box.text_frame.paragraphs[0]
        ru = p_u.add_run()
        ru.text = url_text
        ru.font.name = FONT_MONO
        ru.font.size = Pt(9.5)
        ru.font.color.rgb = SLATE_LIGHT

        # Ảnh thật Retina
        img_path = assets_dir / img_name
        pic = slide.shapes.add_picture(
            str(img_path),
            Inches(left + 0.02),
            Inches(top + bar_h),
            width=Inches(width - 0.04),
            height=Inches(img_h),
        )
        ids.append(pic.shape_id)

        if badge_text:
            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left + width - 2.65),
                Inches(top - 0.18),
                Inches(2.5),
                Inches(0.34),
            )
            badge.adjustments[0] = 0.4
            badge.fill.solid()
            badge.fill.fore_color.rgb = RGBColor(0x06, 0x22, 0x1E)
            badge.line.color.rgb = MINT
            badge.line.width = Pt(1.2)
            p_b = badge.text_frame.paragraphs[0]
            p_b.alignment = PP_ALIGN.CENTER
            rb = p_b.add_run()
            rb.text = f"● {badge_text}"
            rb.font.name = FONT_HEAD
            rb.font.size = Pt(10)
            rb.font.bold = True
            rb.font.color.rgb = WHITE
            ids.append(badge.shape_id)

        return ids

    def attach_transition_and_animations(slide, shape_ids: list[int], transition_kind: str) -> None:
        # 1. Slide Transition (<p:transition>) - đặt trước <p:timing>
        t_el = etree.SubElement(slide._element, f"{{{P_NS}}}transition", spd="med", advClick="1")
        if transition_kind in ("push", "cover"):
            etree.SubElement(t_el, f"{{{P_NS}}}{transition_kind}", dir="r")
        else:
            etree.SubElement(t_el, f"{{{P_NS}}}fade", thruBlk="0")

        if not shape_ids:
            return

        # 2. Staggered Entrance Animations (<p:timing>)
        timing = etree.SubElement(slide._element, f"{{{P_NS}}}timing")
        tn_lst = etree.SubElement(timing, f"{{{P_NS}}}tnLst")
        root_par = etree.SubElement(tn_lst, f"{{{P_NS}}}par")
        root_ctn = etree.SubElement(
            root_par, f"{{{P_NS}}}cTn", id="1", dur="indefinite", restart="never", nodeType="tmRoot"
        )
        root_child = etree.SubElement(root_ctn, f"{{{P_NS}}}childTnLst")
        seq = etree.SubElement(root_child, f"{{{P_NS}}}seq", concurrent="1", nextAc="seek")
        main_ctn = etree.SubElement(
            seq, f"{{{P_NS}}}cTn", id="2", dur="indefinite", nodeType="mainSeq"
        )
        main_child = etree.SubElement(main_ctn, f"{{{P_NS}}}childTnLst")

        next_id = 3
        for idx, spid in enumerate(shape_ids):
            delay_ms = str(idx * 110)

            p1 = etree.SubElement(main_child, f"{{{P_NS}}}par")
            c1 = etree.SubElement(p1, f"{{{P_NS}}}cTn", id=str(next_id), fill="hold")
            next_id += 1
            st1 = etree.SubElement(c1, f"{{{P_NS}}}stCondLst")
            etree.SubElement(st1, f"{{{P_NS}}}cond", delay=delay_ms)
            ch1 = etree.SubElement(c1, f"{{{P_NS}}}childTnLst")

            p2 = etree.SubElement(ch1, f"{{{P_NS}}}par")
            c2 = etree.SubElement(p2, f"{{{P_NS}}}cTn", id=str(next_id), fill="hold")
            next_id += 1
            st2 = etree.SubElement(c2, f"{{{P_NS}}}stCondLst")
            etree.SubElement(st2, f"{{{P_NS}}}cond", delay="0")
            ch2 = etree.SubElement(c2, f"{{{P_NS}}}childTnLst")

            p3 = etree.SubElement(ch2, f"{{{P_NS}}}par")
            c3 = etree.SubElement(
                p3,
                f"{{{P_NS}}}cTn",
                id=str(next_id),
                presetID="10",
                presetClass="entr",
                presetSubtype="0",
                fill="hold",
                nodeType="withEffect",
            )
            next_id += 1
            st3 = etree.SubElement(c3, f"{{{P_NS}}}stCondLst")
            etree.SubElement(st3, f"{{{P_NS}}}cond", delay="0")
            ch3 = etree.SubElement(c3, f"{{{P_NS}}}childTnLst")

            # Set visible
            s_el = etree.SubElement(ch3, f"{{{P_NS}}}set")
            s_bh = etree.SubElement(s_el, f"{{{P_NS}}}cBhvr")
            s_ctn = etree.SubElement(s_bh, f"{{{P_NS}}}cTn", id=str(next_id), dur="1", fill="hold")
            next_id += 1
            etree.SubElement(etree.SubElement(s_ctn, f"{{{P_NS}}}stCondLst"), f"{{{P_NS}}}cond", delay="0")
            etree.SubElement(etree.SubElement(s_bh, f"{{{P_NS}}}tgtEl"), f"{{{P_NS}}}spTgt", spid=str(spid))
            etree.SubElement(etree.SubElement(s_bh, f"{{{P_NS}}}attrNameLst"), f"{{{P_NS}}}attrName").text = "style.visibility"
            etree.SubElement(etree.SubElement(s_el, f"{{{P_NS}}}to"), f"{{{P_NS}}}strVal", val="visible")

            # Fade in
            eff = etree.SubElement(ch3, f"{{{P_NS}}}animEffect", transition="in", filter="fade")
            e_bh = etree.SubElement(eff, f"{{{P_NS}}}cBhvr")
            etree.SubElement(e_bh, f"{{{P_NS}}}cTn", id=str(next_id), dur="420")
            next_id += 1
            etree.SubElement(etree.SubElement(e_bh, f"{{{P_NS}}}tgtEl"), f"{{{P_NS}}}spTgt", spid=str(spid))

            # Subtle Float Up (ppt_y)
            anm = etree.SubElement(ch3, f"{{{P_NS}}}anim", calcmode="lin", valueType="num")
            a_bh = etree.SubElement(anm, f"{{{P_NS}}}cBhvr", additive="base")
            etree.SubElement(a_bh, f"{{{P_NS}}}cTn", id=str(next_id), dur="420", fill="hold")
            next_id += 1
            etree.SubElement(etree.SubElement(a_bh, f"{{{P_NS}}}tgtEl"), f"{{{P_NS}}}spTgt", spid=str(spid))
            etree.SubElement(etree.SubElement(a_bh, f"{{{P_NS}}}attrNameLst"), f"{{{P_NS}}}attrName").text = "ppt_y"
            tav_lst = etree.SubElement(anm, f"{{{P_NS}}}tavLst")
            t0 = etree.SubElement(tav_lst, f"{{{P_NS}}}tav", tm="0")
            etree.SubElement(etree.SubElement(t0, f"{{{P_NS}}}val"), f"{{{P_NS}}}strVal", val="#ppt_y+0.035")
            t1 = etree.SubElement(tav_lst, f"{{{P_NS}}}tav", tm="100000")
            etree.SubElement(etree.SubElement(t1, f"{{{P_NS}}}val"), f"{{{P_NS}}}strVal", val="#ppt_y")

        prev_c = etree.SubElement(seq, f"{{{P_NS}}}prevCondLst")
        pc = etree.SubElement(prev_c, f"{{{P_NS}}}cond", evt="onPrev", delay="0")
        etree.SubElement(etree.SubElement(pc, f"{{{P_NS}}}tgtEl"), f"{{{P_NS}}}sldTgt")

        next_c = etree.SubElement(seq, f"{{{P_NS}}}nextCondLst")
        nc = etree.SubElement(next_c, f"{{{P_NS}}}cond", evt="onNext", delay="0")
        etree.SubElement(etree.SubElement(nc, f"{{{P_NS}}}tgtEl"), f"{{{P_NS}}}sldTgt")

    # =========================================================================
    # DỰNG 14 SLIDE NATIVE POWERPOINT
    # =========================================================================

    # SLIDE 1: HERO COVER
    s1 = prs.slides.add_slide(blank)
    set_dark_bg(s1)
    ids1 = add_header(s1, "HN-KS25-CNTT5 · NHÓM 5", "TechRecruit — Nền tảng Tuyển dụng CNTT", 1)
    ids1.append(add_bento_card(
        s1, 0.65, 1.75, 4.65, 1.75,
        "Hệ sinh thái Tuyển dụng IT",
        ["Kết nối Ứng viên, Nhà tuyển dụng và Quản trị viên với chuẩn hóa kỹ năng, chấm điểm khớp và xác minh pháp nhân."],
        tag="Vanilla JS · FastAPI · MySQL 8.0",
    ))
    ids1.append(add_bento_card(s1, 0.65, 3.7, 1.45, 1.4, "Trang HTML", [], metric="48", highlight=True))
    ids1.append(add_bento_card(s1, 2.25, 3.7, 1.45, 1.4, "API Routers", [], metric="50+", highlight=True))
    ids1.append(add_bento_card(s1, 3.85, 3.7, 1.45, 1.4, "Phân quyền", [], metric="4 Cấp", highlight=True))
    ids1.extend(add_browser_mockup(s1, 5.6, 1.85, 7.05, "01-home.png", "127.0.0.1:5500/frontend/index.html", "4 Cấp Phân quyền RBAC"))
    attach_transition_and_animations(s1, ids1, "fade")

    # SLIDE 2: BỐI CẢNH & BÀI TOÁN
    s2 = prs.slides.add_slide(blank)
    set_dark_bg(s2)
    ids2 = add_header(s2, "BỐI CẢNH & BÀI TOÁN", "Ba Nỗi Đau Tuyển Dụng IT Truyền Thống", 2)
    ids2.append(add_bento_card(
        s2, 0.65, 1.8, 3.8, 4.75, "Với Ứng viên IT",
        ["Tìm kiếm bỏ sót do tên kỹ năng lệch chuẩn (React vs ReactJS).", "Mù mờ lương Gross–Net thực nhận và trạng thái hồ sơ sau khi nộp."],
        tag="Thiếu minh bạch",
    ))
    ids2.append(add_bento_card(
        s2, 4.75, 1.8, 3.8, 4.75, "Với Nhà tuyển dụng",
        ["Khó đánh giá nhanh mức độ phù hợp của hàng chục CV ứng tuyển.", "Dễ đặt trùng lịch phỏng vấn hoặc lộ tệp CV qua URL tĩnh."],
        tag="Vận hành rời rạc",
    ))
    ids2.append(add_bento_card(
        s2, 8.85, 1.8, 3.8, 4.75, "Với Quản trị viên",
        ["Khó chặn doanh nghiệp giả mạo nếu thiếu quy trình duyệt MST & GPKD.", "Thiếu nhật ký kiểm toán (audit_logs) khi xử lý logo và CV."],
        tag="Rủi ro pháp nhân",
    ))
    attach_transition_and_animations(s2, ids2, "push")

    # SLIDE 3: KIẾN TRÚC 3 TẦNG & RBAC
    s3 = prs.slides.add_slide(blank)
    set_dark_bg(s3)
    ids3 = add_header(s3, "KIẾN TRÚC & PHÂN QUYỀN", "Kiến Trúc 3 Tầng & RBAC 4 Cấp Độ", 3)
    ids3.append(add_bento_card(
        s3, 0.65, 1.8, 3.8, 4.75, "1. Frontend · Vanilla JS ES Modules",
        ["48 trang HTML chuẩn hóa route qua routes.js.", "Design Tokens Teal (#0F766E, Be Vietnam Pro).", "Axios interceptor tự động gắn JWT & xử lý 401."],
        tag="check-assets.mjs · 0 lỗi 404",
    ))
    ids3.append(add_bento_card(
        s3, 4.75, 1.8, 3.8, 4.75, "2. Backend API · FastAPI",
        ["50+ router phân tách theo vai trò & nghiệp vụ.", "Thuật toán matching.py & tính lương Gross–Net.", "Kiểm tra Magic Bytes ảnh/PDF & xóa metadata EXIF."],
        tag="Port 8000 · Swagger OpenAPI",
    ))
    ids3.append(add_bento_card(
        s3, 8.85, 1.8, 3.8, 4.75, "3. MySQL 8.0 & Private Vault",
        ["Chuẩn hóa skill_aliases & audit_logs.", "Khóa lạc quan version chống duyệt trùng.", "Lưu CV & giấy phép ngoài public, cấp qua Blob URL."],
        tag="Candidate · Employer · Operator · Admin",
    ))
    attach_transition_and_animations(s3, ids3, "cover")

    # Helper cho các slide Split 40/60 có ảnh Mockup
    def add_split_slide(num, pill, title, cards, img_name, url_text, badge_text, trans="push", ratio_16_10=False):
        sl = prs.slides.add_slide(blank)
        set_dark_bg(sl)
        ids = add_header(sl, pill, title, num)
        y = 1.75
        for c_title, c_lines, c_metric, c_hl in cards:
            h = 1.48 if len(cards) == 3 else 2.2
            ids.append(add_bento_card(sl, 0.65, y, 4.65, h, c_title, c_lines, metric=c_metric, highlight=c_hl))
            y += h + 0.16
        ids.extend(add_browser_mockup(sl, 5.6, 1.85, 7.05, img_name, url_text, badge_text, aspect_16_10=ratio_16_10))
        attach_transition_and_animations(sl, ids, trans)

    # SLIDE 4: XÁC THỰC & ONBOARDING
    add_split_slide(
        4, "PHÂN HỆ 1 · XÁC THỰC & ONBOARDING", "Đăng Ký Theo Vai Trò & Bảo Mật JWT",
        [
            ("Đăng ký phân loại vai trò", ["Chọn trực tiếp Ứng viên hoặc Nhà tuyển dụng ngay khi tạo tài khoản."], "", False),
            ("Bảo mật phiên & Quên mật khẩu", ["Băm bcrypt (12 rounds), CAPTCHA và tăng token_version thu hồi JWT tức thì."], "", False),
            ("Route Guard tự động", ["Điều hướng đúng không gian làm việc hoặc chặn về 403.html khi sai quyền."], "", False),
        ],
        "02-login.png", "127.0.0.1:5500/frontend/pages/login.html", "JWT + token_version", "fade"
    )

    # SLIDE 5: ỨNG VIÊN — TÌM VIỆC
    add_split_slide(
        5, "PHÂN HỆ 2 · TRẢI NGHIỆM ỨNG VIÊN", "Tìm Việc Đa Tiêu Chí & Ứng Tuyển Một Chạm",
        [
            ("Đang mở tuyển với Facets", ["Bộ lọc thời gian thực theo địa điểm, cấp bậc và công ty."], "58 Việc làm", True),
            ("Lọc thông minh & Chi tiết JD", ["Chuẩn hóa kỹ năng IT qua bảng skill_aliases."], "", False),
            ("Nộp hồ sơ nhanh gọn", ["Chọn CV đã lưu hoặc tải lên kèm Cover Letter và thông báo xác nhận."], "", False),
        ],
        "02-jobs.png", "127.0.0.1:5500/frontend/pages/jobs.html", "Facets Thời gian thực", "push"
    )

    # SLIDE 6: ỨNG VIÊN — TẠO CV & THEO DÕI
    add_split_slide(
        6, "PHÂN HỆ 2 · TRÌNH TẠO CV & THEO DÕI", "Mẫu CV Chuẩn A4 & Theo Dõi Trạng Thái Thực",
        [
            ("Trình tạo CV tương tác", ["4 phong cách trình bày (cv-builder.html), xuất PDF chuẩn A4."], "", False),
            ("Lưu tin & Lộ trình nghề nghiệp", ["Lưu việc quan tâm, phân tích kỹ năng thiếu và tính lương Gross–Net."], "", False),
            ("Tiến trình xét duyệt minh bạch", ["Theo dõi statusTimeline từ nộp đơn đến phỏng vấn, rút hồ sơ chủ động."], "", False),
        ],
        "05-cv-templates.png", "127.0.0.1:5500/frontend/pages/cv-templates.html", "Xuất PDF Chuẩn A4", "push"
    )

    # SLIDE 7: NHÀ TUYỂN DỤNG — CÔNG TY & ĐĂNG TIN
    add_split_slide(
        7, "PHÂN HỆ 3 · NHÀ TUYỂN DỤNG", "Hồ Sơ Doanh Nghiệp & Đăng Tin Live Preview",
        [
            ("Thiết lập công ty & Gửi MST", ["Cập nhật logo, nộp MST + GPKD để mở khóa quyền đăng tin."], "", False),
            ("Đăng tin có Live Preview", ["Biểu mẫu job-form.html hiển thị thẻ xem trước trực quan khi soạn thảo."], "", False),
            ("Chốt chặn Lưu nháp an toàn", ["Công ty chưa xác minh vẫn lưu draft, chặn xuất bản với mã 403."], "", False),
        ],
        "08-hr-create-job.png", "127.0.0.1:5500/frontend/pages/employer/job-form.html", "Live Preview Card", "push", ratio_16_10=True
    )

    # SLIDE 8: NHÀ TUYỂN DỤNG — QUẢN LÝ TIN & SÀNG LỌC
    s8 = prs.slides.add_slide(blank)
    set_dark_bg(s8)
    ids8 = add_header(s8, "PHÂN HỆ 3 · SÀNG LỌC ỨNG VIÊN", "Quản Lý Trạng Thái Tin & Drawer Hồ Sơ Bảo Mật", 8)
    ids8.append(add_bento_card(
        s8, 0.65, 1.8, 3.8, 4.75, "Vòng đời Tin tuyển dụng",
        ["Quản lý tập trung trạng thái: Nháp (draft), Chờ duyệt, Đang tuyển và Chờ xác minh lại (reverify_pending)."],
        tag="employerJobs.js",
    ))
    ids8.append(add_bento_card(
        s8, 4.75, 1.8, 3.8, 4.75, "Drawer & Matching Score",
        ["Ngăn kéo employerCandidateDrawer.js hiển thị điểm độ khớp kỹ năng (%), Cover Letter và ghi chú nội bộ."],
        tag="Khớp 85% Kỹ năng",
    ))
    ids8.append(add_bento_card(
        s8, 8.85, 1.8, 3.8, 4.75, "Mở CV bằng Blob URL",
        ["Tải luồng nhị phân qua Bearer JWT, nhúng PDF trong iframe và gọi URL.revokeObjectURL khi đóng."],
        tag="Zero Token Leak",
    ))
    attach_transition_and_animations(s8, ids8, "fade")

    # SLIDE 9: NHÀ TUYỂN DỤNG — ĐIỀU PHỐI PHỎNG VẤN
    s9 = prs.slides.add_slide(blank)
    set_dark_bg(s9)
    ids9 = add_header(s9, "PHÂN HỆ 3 · ĐIỀU PHỐI PHỎNG VẤN", "Lịch Phỏng Vấn Chống Trùng & Cập Nhật Vòng Tuyển", 9)
    ids9.append(add_bento_card(
        s9, 0.65, 1.8, 7.1, 4.75, "Cơ chế Chống trùng lịch (409 Conflict)",
        [
            "Kiểm tra giao nhau khung giờ phỏng vấn trên MySQL trước khi tạo hoặc dời lịch hẹn.",
            "Hỗ trợ cả phỏng vấn Online (Meeting Link) và Offline (địa chỉ văn phòng).",
            "Tự động bắn thông báo kèm email nhắc lịch tới ứng viên ngay khi xác nhận.",
        ],
        tag="employer_interviews.py · employerInterviewCalendar.js",
    ))
    ids9.append(add_bento_card(
        s9, 8.05, 1.8, 4.6, 2.25, "Chặn trùng khung giờ",
        ["Chặn tuyệt đối hai lịch phỏng vấn trùng khung giờ của cùng người phụ trách."],
        metric="409 Conflict", highlight=True,
    ))
    ids9.append(add_bento_card(
        s9, 8.05, 4.3, 4.6, 2.25, "Cập nhật kết quả vòng tuyển",
        ["Chuyển trạng thái hồ sơ (Mời phỏng vấn → Đạt → Offer) đồng bộ tức thì sang Ứng viên."],
    ))
    attach_transition_and_animations(s9, ids9, "push")

    # SLIDE 10: QUẢN TRỊ VIÊN — DASHBOARD & TÀI KHOẢN
    s10 = prs.slides.add_slide(blank)
    set_dark_bg(s10)
    ids10 = add_header(s10, "PHÂN HỆ 4 · QUẢN TRỊ VIÊN", "Bảng Điều Khiển & Quản Lý Tài Khoản Hệ Thống", 10)
    ids10.append(add_bento_card(
        s10, 0.65, 1.8, 3.8, 4.75, "Layout Dùng chung",
        ["adminLayout.js đồng bộ trạng thái thu gọn sidebar (256px ↔ 72px) ngay từ <head>, hỗ trợ tooltip và drawer."],
        tag="Zero FOUC",
    ))
    ids10.append(add_bento_card(
        s10, 4.75, 1.8, 3.8, 4.75, "Quản lý Tài khoản",
        ["Tìm kiếm, lọc theo vai trò (candidate, employer, operator, admin) và trạng thái hoạt động."],
        tag="pages/admin/users.html",
    ))
    ids10.append(add_bento_card(
        s10, 8.85, 1.8, 3.8, 4.75, "Khóa / Mở khóa An toàn",
        ["Cập nhật is_active kèm lý do khóa (locked_reason) và vô hiệu hóa phiên đăng nhập hiện tại."],
        tag="users.py",
    ))
    attach_transition_and_animations(s10, ids10, "fade")

    # SLIDE 11: QUẢN TRỊ VIÊN — XÁC MINH PHÁP NHÂN
    add_split_slide(
        11, "PHÂN HỆ 4 · XÁC MINH PHÁP NHÂN", "Thẩm Định Doanh Nghiệp & Khóa Lạc Quan",
        [
            ("Chuẩn hóa MST & Magic Bytes", ["Kiểm tra MST 10/13 số; chứng từ PNG/JPEG/PDF ≤ 5MB lưu trong company_licenses/."], "", False),
            ("Chống duyệt đè (Optimistic Lock)", ["Kiểm tra version + Idempotency-Key khi duyệt/từ chối/thu hồi, phát đúng 01 thông báo."], "", False),
            ("Ghi vết Audit Log đầy đủ", ["Mỗi lần mở xem giấy phép (view_license) hay từ chối/thu hồi đều lưu vào audit_logs."], "", False),
        ],
        "11-admin-verify.png", "127.0.0.1:5500/frontend/pages/admin/companies.html", "Đã xác minh MST", "push", ratio_16_10=True
    )

    # SLIDE 12: QUẢN TRỊ VIÊN — KIỂM DUYỆT TIN, LOGO & KHO CV
    add_split_slide(
        12, "PHÂN HỆ 4 · KIỂM DUYỆT & CÔNG CỤ", "Kiểm Duyệt Tin, Logo WebP & Kho CV Toàn Hệ Thống",
        [
            ("Chuẩn hóa Logo PNG/WebP", ["Kiểm tra chữ ký nhị phân, tái tạo ảnh bằng Pillow xóa 100% metadata EXIF/ICC."], "", False),
            ("Duyệt tin & Kho CV Admin", ["Duyệt/gỡ tin vi phạm (admin_jobs.py) và xem CV toàn hệ thống có ghi nhật ký kiểm toán."], "", False),
            ("Hệ công cụ tài chính IT", ["Tính lương Gross–Net (Gross 30M → Net 25.22M ₫) minh bạch BHXH & thuế TNCN."], "", False),
        ],
        "04-companies.png", "127.0.0.1:5500/frontend/pages/companies.html", "Logo WebP Chuẩn hóa", "push"
    )

    # SLIDE 13: ĐIỂM NHẤN KỸ THUẬT & BẢO MẬT
    s13 = prs.slides.add_slide(blank)
    set_dark_bg(s13)
    ids13 = add_header(s13, "ĐIỂM NHẤN KỸ THUẬT & BẢO MẬT", "Chốt Chặn An Toàn & Kiểm Thử Tự Động", 13)
    ids13.append(add_bento_card(
        s13, 0.65, 1.8, 3.8, 4.75, "Thu hồi JWT tức thì",
        ["Mỗi lần đăng xuất hoặc đổi mật khẩu tăng token_version += 1, vô hiệu hóa toàn bộ token cũ ngay lập tức."],
        tag="role_required() · 4 cấp độ",
    ))
    ids13.append(add_bento_card(
        s13, 4.75, 1.8, 3.8, 4.75, "Zero Trust Upload",
        ["Xác thực Magic Bytes chặn file giả mạo đuôi, tách thư mục riêng tư và khóa duy nhất pending_company_lock."],
        tag="Optimistic Lock · Audit Log",
    ))
    ids13.append(add_bento_card(
        s13, 8.85, 1.8, 3.8, 4.75, "Kiểm thử toàn diện",
        ["Hơn 50 tệp kiểm thử pytest cho API/nghiệp vụ kết hợp node:test và check-assets.mjs."],
        tag="48 trang HTML · 0 broken links",
    ))
    attach_transition_and_animations(s13, ids13, "fade")

    # SLIDE 14: LỘ TRÌNH & KẾT LUẬN
    s14 = prs.slides.add_slide(blank)
    set_dark_bg(s14)
    ids14 = add_header(s14, "LỘ TRÌNH & KẾT LUẬN", "Hướng Mở Rộng & Lời Cảm Ơn", 14)
    ids14.append(add_bento_card(
        s14, 0.65, 1.8, 3.8, 4.75, "AI Matching Ngữ nghĩa",
        ["Kết hợp Vector Embedding để gợi ý việc làm và phân tích CV tự động chính xác hơn theo ngữ cảnh."],
        tag="Phase 2 · AI",
    ))
    ids14.append(add_bento_card(
        s14, 4.75, 1.8, 3.8, 4.75, "Chat & Realtime WebSocket",
        ["Trao đổi trực tiếp giữa Ứng viên và HR ngay trong Drawer kèm thông báo tức thì đa thiết bị."],
        tag="Phase 2 · Realtime",
    ))
    ids14.append(add_bento_card(
        s14, 8.85, 1.8, 3.8, 4.75, "Nhóm 5 · HN-KS25-CNTT5",
        ["Trân trọng cảm ơn Quý Thầy/Cô và Các bạn đã theo dõi. Nhóm sẵn sàng Demo & Q&A."],
        tag="TechRecruit 2026", metric="Cảm ơn!", highlight=True,
    ))
    attach_transition_and_animations(s14, ids14, "fade")

    prs.save(str(output_pptx))
    print(f"Xuất thành công Native PowerPoint (14 slides có Animation từng thẻ): {output_pptx}")


if __name__ == "__main__":
    main()
