#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
سكربت سحب بيانات المنتجات العامة من متجر mokab.com (منصة سلة - Salla)
=====================================================================

المنهجية (تم التحقق منها فعليًا قبل كتابة هذا السكربت):

1) الموقع مبني على منصة سلة. صفحات المتجر تستخدم واجهة سلة العامة
   (api.salla.dev) لعرض بيانات المنتجات في المتصفح. هذه الواجهة عامة
   تمامًا ولا تحتاج تسجيل دخول أو كوكيز - فقط بضعة Headers ثابتة
   (تم اكتشافها بمراقبة الطلبات الفعلية من متصفح حقيقي على الموقع).

2) نقطة /store/v1/products/{id}/details تُرجع بشكل موثوق لكل معرف
   منتج: الاسم، الوصف الكامل، السعر الحالي، سعر ما قبل الخصم، السعر
   بعد الخصم، SKU/رقم الموديل، الضمان (يظهر كحقل subtitle)، الكمية،
   حالة التوفر، ورابط المنتج.

3) نقطة قوائم التصنيفات (source=category) في تجربتنا كانت غير موثوقة
   (تُرجع نفس المنتجات بغض النظر عن التصنيف المطلوب - على الأرجح
   بسبب طبقة تخزين مؤقت/كاش لدى سلة). لذلك لا نعتمد عليها كمصدر
   وحيد. بدلاً من ذلك نعتمد على:
      - sitemap-2.xml : يحتوي فعليًا على قائمة كاملة بروابط كل
        المنتجات المفهرسة في الموقع (وجدنا فيه نحو 4999 رابط منتج
        مطابق لنمط /pNNNNNNNN عند فحصه مباشرة - وهو أدق مصدر عام
        متاح لعدد المنتجات الكلي).
      - صفحة كل منتج (HTML) لاستخراج: التصنيف (من Meta Tag
        product:category وهو مطابق لمسار التصنيفات الحقيقي)،
        والعلامة التجارية (من بيانات Schema.org JSON-LD المضمّنة، وهي
        نفس البيانات التي تقرأها جوجل)، وأي خيارات ألوان/مقاسات ظاهرة
        كعناصر اختيار داخل نفس المنتج (ملاحظة: أغلب المتغيرات في هذا
        المتجر تُباع كمنتجات مستقلة لكل لون/مقاس، وليس كخيارات داخل المنتج).

المخرجات: افتراضيًا يبني ملف "mokab_dashboard.html" - لوحة تحكم واحدة
ذاتية الاحتواء (الصور والبيانات مضمّنة داخل الملف نفسه، تفتح بأي متصفح
بدون إنترنت لعرض البيانات - تحتاج فقط إنترنت لعرض صور المنتجات نفسها
لأنها تُحمّل من سيرفر مكعب) فيها: تحليل عام (إحصاءات، أعلى التصنيفات
والعلامات التجارية)، ثم كل المنتجات كبطاقات (صورة أمام كل منتج) أو
كجدول قابل للفرز والفلترة والبحث. يمكن أيضًا اختيار Excel/CSV/JSON
عبر --format.

الاستخدام:
    pip install requests beautifulsoup4 lxml pandas openpyxl --break-system-packages

    # التشغيل الافتراضي: يسحب كل منتجات المتجر (~5000) وينتج لوحة تحكم HTML:
    python mokab_scraper.py

    # تجربة سريعة على 30 منتج فقط:
    python mokab_scraper.py --limit 30

    # إخراج Excel بدل لوحة التحكم:
    python mokab_scraper.py --format xlsx

ملاحظات أخلاقية/قانونية:
- هذا السكربت يسحب فقط البيانات العامة الظاهرة للزوار (بدون تسجيل
  دخول، بدون تجاوز أي حماية). يُنصح بإبقاء التأخير والتزامن معقولين
  حتى لا يُحمّل الخادم فوق طاقته، ومراجعة شروط استخدام الموقع قبل
  أي استخدام تجاري للبيانات.
"""

import argparse
import concurrent.futures
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field, asdict
from typing import Optional

import requests
from bs4 import BeautifulSoup

# ============================================================================
# إعدادات سريعة لمن لا يستخدم سطر الأوامر (تشغيل مباشر من IDLE بالزر F5)
# عدّل القيم هنا فقط ثم احفظ الملف (Ctrl+S) وشغّله - لا حاجة لأي كتابة أوامر.
# ============================================================================
DEFAULT_LIMIT = None        # None = سحب كل منتجات المتجر (~5000). ضع رقمًا (مثلاً 20) لتجربة سريعة
DEFAULT_FORMAT = "dashboard"  # "dashboard" (لوحة تحكم كاملة بالصور والتحليل) أو "xlsx" أو "csv" أو "json" أو "html"
DEFAULT_WORKERS = 8         # عدد الاتصالات المتزامنة (8 قيمة معقولة وآمنة)
DEFAULT_DOWNLOAD_IMAGES = False  # اجعلها True لتحميل ملفات الصور فعليًا وليس فقط روابطها
# ============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

STORE_IDENTIFIER = "752550455"  # معرف متجر مكعب على سلة (عام وظاهر في روابط الصور CDN الخاصة بالمتجر)
API_BASE = "https://api.salla.dev/store/v1"
SITE_BASE = "https://mokab.com"

HEADERS_API = {
    "Accept": "application/json",
    "Store-Identifier": STORE_IDENTIFIER,
    "S-SOURCE": "twilight",
    "S-APP-VERSION": "2.14.529",
    "S-APP-OS": "browser",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/122.0 Safari/537.36",
}

HEADERS_HTML = {
    "User-Agent": HEADERS_API["User-Agent"],
    "Accept-Language": "ar,en;q=0.8",
}

REQUEST_TIMEOUT = 20
RETRY_COUNT = 3
RETRY_BACKOFF = 2.0
POLITE_DELAY = 0.15  # ثانية بين الطلبات لكل عامل (Thread) لتخفيف الضغط على الخادم


@dataclass
class Product:
    id: str = ""
    name: str = ""
    price_current: Optional[float] = None
    price_before_discount: Optional[float] = None
    discount_percent: Optional[float] = None
    category: str = ""
    brand: str = ""
    description: str = ""
    sku: str = ""
    options: str = ""
    availability: str = ""
    warranty: str = ""
    specs_text: str = ""
    url: str = ""
    quantity: Optional[float] = None
    image_urls: str = ""
    in_stock: Optional[bool] = None
    rating_value: Optional[float] = None
    review_count: Optional[int] = None
    errors: str = ""


def _get_with_retry(url, headers, session):
    last_exc = None
    for attempt in range(RETRY_COUNT):
        try:
            resp = session.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            if resp.status_code == 200:
                return resp
            if resp.status_code == 429:
                time.sleep(RETRY_BACKOFF * (attempt + 1) * 3)
                continue
            last_exc = RuntimeError(f"HTTP {resp.status_code} for {url}")
        except requests.RequestException as e:
            last_exc = e
        time.sleep(RETRY_BACKOFF * (attempt + 1))
    raise last_exc if last_exc else RuntimeError(f"Failed to fetch {url}")


def get_all_product_urls(session, category_filter=None, limit=None):
    """يجمع كل روابط المنتجات من ملفات خريطة الموقع (sitemap)."""
    index_resp = _get_with_retry(f"{SITE_BASE}/sitemap.xml", HEADERS_HTML, session)
    soup = BeautifulSoup(index_resp.text, "xml")
    sub_sitemaps = [loc.text for loc in soup.find_all("loc")]

    product_urls = []
    for sm_url in sub_sitemaps:
        if "blog" in sm_url:
            continue
        resp = _get_with_retry(sm_url, HEADERS_HTML, session)
        sm_soup = BeautifulSoup(resp.text, "xml")
        for loc in sm_soup.find_all("loc"):
            url = loc.text.strip()
            m = re.search(r"/p(\d+)(?:$|[/?#])", url)
            if m:
                product_urls.append((m.group(1), url))

    # إزالة التكرار مع الحفاظ على الترتيب
    seen = set()
    unique = []
    for pid, url in product_urls:
        if pid not in seen:
            seen.add(pid)
            unique.append((pid, url))

    if limit:
        unique = unique[:limit]

    return unique


def parse_html_extras(html: str):
    """يستخرج من صفحة المنتج (HTML) الحقول غير المتوفرة عبر الـ API:
    التصنيف، العلامة التجارية، وأي خيارات ألوان/مقاسات ظاهرة."""
    soup = BeautifulSoup(html, "lxml")
    extras = {
        "category": "", "brand": "", "options": "", "availability_meta": "",
        "images": [], "in_stock": None, "rating_value": None, "review_count": None,
    }

    # مصدر إضافي موثوق جدًا: بيانات Schema.org (JSON-LD) المضمّنة في كل صفحة منتج
    # (نفس البيانات التي تقرأها جوجل) - تعطي العلامة التجارية الحقيقية، حالة
    # التوفر بصيغة معيارية، ومتوسط التقييم وعدد المراجعات إن وُجدت.
    for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            payload = json.loads(script.string or "{}")
        except Exception:
            continue
        graph = payload.get("@graph", [payload])
        for node in graph:
            if node.get("@type") != "Product":
                continue
            brand = node.get("brand")
            if isinstance(brand, dict) and brand.get("name"):
                extras["brand"] = brand["name"]
            offers = node.get("offers")
            if isinstance(offers, dict):
                avail = (offers.get("availability") or "").lower()
                if "instock" in avail:
                    extras["in_stock"] = True
                elif "outofstock" in avail or "soldout" in avail:
                    extras["in_stock"] = False
            agg = node.get("aggregateRating")
            if isinstance(agg, dict):
                try:
                    extras["rating_value"] = float(agg.get("ratingValue"))
                except (TypeError, ValueError):
                    pass
                try:
                    extras["review_count"] = int(agg.get("reviewCount") or agg.get("ratingCount"))
                except (TypeError, ValueError):
                    pass
            reviews = node.get("review")
            if isinstance(reviews, list) and extras["review_count"] is None:
                extras["review_count"] = len(reviews)

    # معرض صور المنتج: سلة تضع كل صورة أصلية داخل رابط
    # <a data-fslightbox="product_<id>" href="...">  وهذا موجود في الـ HTML
    # الأساسي (بدون الحاجة لتشغيل جافاسكربت)
    gallery_links = soup.select("a[data-fslightbox]")
    seen_imgs = set()
    for a in gallery_links:
        href = a.get("href")
        if href and href not in seen_imgs:
            seen_imgs.add(href)
            extras["images"].append(href)
    if not extras["images"]:
        # نسخة احتياطية: أخذ og:image من الميتا تاق
        og_img = soup.find("meta", attrs={"property": "og:image"})
        if og_img and og_img.get("content"):
            extras["images"].append(og_img["content"])

    cat_tag = soup.find("meta", attrs={"property": "product:category"})
    if cat_tag and cat_tag.get("content"):
        extras["category"] = " > ".join(
            [c.strip() for c in cat_tag["content"].split(",") if c.strip()]
        )

    avail_tag = soup.find("meta", attrs={"property": "product:availability"})
    if avail_tag and avail_tag.get("content"):
        extras["availability_meta"] = avail_tag["content"].strip()

    # احتياط فقط إن لم يوجد JSON-LD: رابط العلامة التجارية أعلى صفحة المنتج (/brand-<id>)
    if not extras["brand"]:
        brand_link = soup.find("a", href=re.compile(r"/brand-\d+"))
        if brand_link:
            text = brand_link.get_text(strip=True)
            if text and "المزيد من" not in text:
                extras["brand"] = text
            else:
                img = brand_link.find("img")
                if img and img.get("alt"):
                    extras["brand"] = img["alt"].strip()

    # محاولة أفضل جهد لاكتشاف خيارات لون/مقاس داخل نفس المنتج (نادرة في هذا المتجر
    # لأن أغلب المتغيرات تُباع كمنتجات مستقلة لكل لون/مقاس)
    option_values = set()
    for sel in soup.select("select option"):
        txt = sel.get_text(strip=True)
        if txt and "اختر" not in txt and "select" not in txt.lower():
            option_values.add(txt)
    for btn in soup.select("[data-option-value], .option-value, .s-product-options-list li"):
        txt = btn.get_text(strip=True)
        if txt:
            option_values.add(txt)
    if option_values:
        extras["options"] = "، ".join(sorted(option_values))

    return extras


def download_product_images(session, pid: str, image_urls: list, images_dir: str):
    """يحمّل صور المنتج فعليًا إلى مجلد images/<pid>/ (اختياري عبر --download-images)."""
    saved_paths = []
    product_dir = os.path.join(images_dir, pid)
    os.makedirs(product_dir, exist_ok=True)
    for i, img_url in enumerate(image_urls, start=1):
        ext = os.path.splitext(img_url.split("?")[0])[1] or ".jpg"
        if len(ext) > 5:  # احتياط لو الامتداد غير منطقي
            ext = ".jpg"
        out_file = os.path.join(product_dir, f"{i}{ext}")
        try:
            resp = _get_with_retry(img_url, HEADERS_HTML, session)
            with open(out_file, "wb") as f:
                f.write(resp.content)
            saved_paths.append(out_file)
        except Exception:
            pass
    return saved_paths


def fetch_product(session, pid: str, url: str, images_dir: Optional[str] = None) -> Product:
    p = Product(id=pid, url=url)
    errors = []

    # 1) بيانات دقيقة عبر API سلة العام
    try:
        api_url = f"{API_BASE}/products/{pid}/details"
        resp = _get_with_retry(api_url, HEADERS_API, session)
        data = resp.json().get("data", {})

        p.name = data.get("name", "") or ""
        p.description = re.sub(
            r"\s+", " ", BeautifulSoup(data.get("description") or "", "lxml").get_text(" ")
        ).strip()
        p.sku = data.get("sku", "") or ""
        p.warranty = data.get("subtitle", "") or ""
        p.price_current = data.get("sale_price") or data.get("price")
        p.price_before_discount = data.get("regular_price") or data.get("price")
        if p.price_before_discount and p.price_current and p.price_before_discount > p.price_current:
            p.discount_percent = round(
                (1 - (p.price_current / p.price_before_discount)) * 100, 1
            )
        p.quantity = data.get("quantity")
        if data.get("is_out_of_stock"):
            p.availability = "غير متوفر"
            p.in_stock = False
        elif data.get("is_available"):
            p.availability = "متوفر"
            p.in_stock = True
        else:
            p.availability = "غير محدد"
        if not p.url:
            p.url = data.get("url", "") or url
    except Exception as e:
        errors.append(f"api:{e}")

    # 2) تصنيف + علامة تجارية + خيارات عبر صفحة المنتج نفسها
    try:
        html_resp = _get_with_retry(url, HEADERS_HTML, session)
        extras = parse_html_extras(html_resp.text)
        p.category = extras["category"]
        p.brand = extras["brand"]
        p.options = extras["options"]
        p.image_urls = " | ".join(extras["images"])
        p.rating_value = extras["rating_value"]
        p.review_count = extras["review_count"]
        if p.in_stock is None and extras["in_stock"] is not None:
            p.in_stock = extras["in_stock"]
            p.availability = "متوفر" if extras["in_stock"] else "غير متوفر"
        if not p.availability and extras["availability_meta"]:
            p.availability = extras["availability_meta"]

        if images_dir and extras["images"]:
            download_product_images(session, pid, extras["images"], images_dir)
    except Exception as e:
        errors.append(f"html:{e}")

    p.errors = "; ".join(errors)
    time.sleep(POLITE_DELAY)
    return p


def _esc(text) -> str:
    """تهريب بسيط للنصوص داخل HTML لتفادي كسر الصفحة."""
    if text is None:
        return ""
    text = str(text)
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def build_html_catalog(products, out_path: str):
    """يبني صفحة HTML واحدة (كتالوج) تعرض كل منتج كبطاقة فيها صورته مباشرة،
    مناسبة للتصفح بالمتصفح مباشرة (مثل معاينة سريعة للمتجر بالكامل)."""
    cards = []
    for p in products:
        first_image = (p.image_urls.split(" | ")[0] if p.image_urls else "")
        price_html = ""
        if p.price_current is not None:
            price_html = f'<span class="price-now">{_esc(p.price_current)} ر.س</span>'
            if p.price_before_discount and p.price_before_discount > (p.price_current or 0):
                price_html += f' <span class="price-old">{_esc(p.price_before_discount)} ر.س</span>'
                if p.discount_percent:
                    price_html += f' <span class="badge-discount">خصم {_esc(p.discount_percent)}%</span>'

        avail_class = "in-stock" if p.in_stock else ("out-of-stock" if p.in_stock is False else "")
        avail_label = _esc(p.availability) or "غير محدد"

        img_html = (
            f'<img src="{_esc(first_image)}" alt="{_esc(p.name)}" loading="lazy">'
            if first_image else '<div class="no-image">لا توجد صورة</div>'
        )

        cards.append(f"""
        <a class="card" href="{_esc(p.url)}" target="_blank" rel="noopener">
          <div class="thumb">{img_html}</div>
          <div class="body">
            <div class="badge {avail_class}">{avail_label}</div>
            <h3>{_esc(p.name)}</h3>
            <div class="price">{price_html}</div>
            <div class="meta">{_esc(p.category)}</div>
            <div class="meta">{_esc(p.brand)}</div>
            <div class="meta">SKU: {_esc(p.sku)}</div>
          </div>
        </a>""")

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>كتالوج منتجات مكعب - {len(products)} منتج</title>
<style>
  body {{ font-family: "Segoe UI", Tahoma, Arial, sans-serif; background:#f5f6f8; margin:0; padding:24px; }}
  h1 {{ text-align:center; color:#1e3957; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fill, minmax(220px,1fr)); gap:16px; max-width:1400px; margin:0 auto; }}
  .card {{ background:#fff; border-radius:10px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,.1); text-decoration:none; color:#222; display:flex; flex-direction:column; transition:transform .15s; }}
  .card:hover {{ transform:translateY(-3px); box-shadow:0 4px 10px rgba(0,0,0,.15); }}
  .thumb {{ width:100%; aspect-ratio:1/1; background:#fafafa; display:flex; align-items:center; justify-content:center; }}
  .thumb img {{ width:100%; height:100%; object-fit:contain; }}
  .no-image {{ color:#999; font-size:13px; }}
  .body {{ padding:10px 12px; flex:1; display:flex; flex-direction:column; gap:4px; }}
  h3 {{ font-size:14px; margin:2px 0; line-height:1.4; min-height:40px; }}
  .price-now {{ font-weight:bold; color:#1e3957; font-size:15px; }}
  .price-old {{ color:#999; text-decoration:line-through; font-size:12px; }}
  .badge-discount {{ background:#e74c3c; color:#fff; border-radius:6px; padding:1px 6px; font-size:11px; }}
  .meta {{ font-size:12px; color:#666; }}
  .badge {{ align-self:flex-start; font-size:11px; border-radius:6px; padding:2px 8px; background:#eee; color:#555; }}
  .badge.in-stock {{ background:#e6f7ec; color:#1b8a4a; }}
  .badge.out-of-stock {{ background:#fdeaea; color:#c0392b; }}
</style>
</head>
<body>
<h1>كتالوج منتجات مكعب ({len(products)} منتج)</h1>
<div class="grid">
{"".join(cards)}
</div>
</body>
</html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)


DASHBOARD_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>لوحة تحكم منتجات مكعب</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{
    --claude-orange:#DA7756;
    --claude-orange-dark:#C15F3C;
    --bg:#F7F5F1;
    --panel:#FFFFFF;
    --ink:#3D3929;
    --ink-soft:#6B6555;
    --border:#E8E4DA;
    --green:#1B8A4A;
    --green-bg:#E6F7EC;
    --red:#C0392B;
    --red-bg:#FDEAEA;
  }
  *{box-sizing:border-box;}
  body{
    margin:0; background:var(--bg); color:var(--ink);
    font-family:"IBM Plex Sans Arabic", Tahoma, Arial, sans-serif;
  }
  header{
    background:linear-gradient(135deg, var(--claude-orange), var(--claude-orange-dark));
    color:#fff; padding:28px 24px 60px;
  }
  header .top{ display:flex; align-items:center; justify-content:space-between; max-width:1400px; margin:0 auto; flex-wrap:wrap; gap:10px;}
  header h1{ margin:0; font-size:22px; font-weight:700; }
  header p{ margin:6px 0 0; opacity:.92; font-size:13px; }
  .wrap{ max-width:1400px; margin:-40px auto 40px; padding:0 16px; }

  .stats-grid{ display:grid; grid-template-columns:repeat(auto-fit, minmax(150px,1fr)); gap:12px; margin-bottom:20px; }
  .stat-card{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:16px; text-align:center; box-shadow:0 2px 8px rgba(0,0,0,.04); }
  .stat-card .num{ font-size:24px; font-weight:700; color:var(--claude-orange-dark); }
  .stat-card .label{ font-size:12px; color:var(--ink-soft); margin-top:4px; }

  .panel{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:18px; margin-bottom:18px; box-shadow:0 2px 8px rgba(0,0,0,.04); }
  .panel h2{ font-size:15px; margin:0 0 12px; }
  .bar-row{ display:flex; align-items:center; gap:10px; margin-bottom:8px; font-size:13px; }
  .bar-row .name{ width:170px; flex-shrink:0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:var(--ink-soft); }
  .bar-row .track{ flex:1; background:var(--bg); border-radius:6px; height:10px; overflow:hidden; }
  .bar-row .fill{ height:100%; background:var(--claude-orange); border-radius:6px; }
  .bar-row .count{ width:38px; text-align:left; font-size:12px; color:var(--ink-soft); }
  .analysis-cols{ display:grid; grid-template-columns:1fr 1fr; gap:18px; }
  @media (max-width:800px){ .analysis-cols{ grid-template-columns:1fr; } }

  .controls{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; }
  .controls input[type=text], .controls select{
    font-family:inherit; padding:9px 12px; border-radius:10px; border:1px solid var(--border);
    background:var(--bg); color:var(--ink); font-size:13px;
  }
  .controls input[type=text]{ flex:1; min-width:200px; }
  .view-toggle{ display:flex; border:1px solid var(--border); border-radius:10px; overflow:hidden; }
  .view-toggle button{
    font-family:inherit; border:0; background:var(--bg); color:var(--ink-soft); padding:9px 16px; font-size:13px; cursor:pointer;
  }
  .view-toggle button.active{ background:var(--claude-orange); color:#fff; }
  .result-count{ font-size:13px; color:var(--ink-soft); margin:10px 2px; }

  /* بطاقات */
  .grid{ display:grid; grid-template-columns:repeat(auto-fill, minmax(210px,1fr)); gap:14px; }
  .card{ background:var(--panel); border:1px solid var(--border); border-radius:14px; overflow:hidden; display:flex; flex-direction:column; text-decoration:none; color:var(--ink); transition:transform .15s, box-shadow .15s; }
  .card:hover{ transform:translateY(-3px); box-shadow:0 6px 16px rgba(0,0,0,.1); }
  .thumb{ width:100%; aspect-ratio:1/1; background:#faf9f6; display:flex; align-items:center; justify-content:center; }
  .thumb img{ width:100%; height:100%; object-fit:contain; }
  .no-image{ color:#bbb; font-size:12px; }
  .card .body{ padding:10px 12px 14px; flex:1; display:flex; flex-direction:column; gap:5px; }
  .card h3{ font-size:13px; margin:2px 0; line-height:1.45; min-height:38px; font-weight:500; }
  .price-now{ font-weight:700; color:var(--claude-orange-dark); font-size:15px; }
  .price-old{ color:#a39c86; text-decoration:line-through; font-size:11px; margin-inline-start:6px; }
  .badge-discount{ background:var(--red); color:#fff; border-radius:6px; padding:1px 6px; font-size:10px; margin-inline-start:6px; }
  .meta{ font-size:11.5px; color:var(--ink-soft); }
  .badge{ align-self:flex-start; font-size:10.5px; border-radius:6px; padding:2px 8px; background:var(--bg); color:var(--ink-soft); }
  .badge.in-stock{ background:var(--green-bg); color:var(--green); }
  .badge.out-of-stock{ background:var(--red-bg); color:var(--red); }

  /* جدول */
  table{ width:100%; border-collapse:collapse; font-size:12.5px; }
  thead th{ position:sticky; top:0; background:var(--bg); text-align:right; padding:9px 8px; border-bottom:2px solid var(--border); white-space:nowrap; }
  tbody td{ padding:8px; border-bottom:1px solid var(--border); vertical-align:middle; }
  tbody tr:hover{ background:#FBF9F5; }
  .t-thumb{ width:46px; height:46px; object-fit:contain; border-radius:6px; background:#faf9f6; }
  .t-name{ max-width:280px; }
  .t-link{ color:var(--claude-orange-dark); text-decoration:none; font-weight:500; }

  .pager{ display:flex; justify-content:center; align-items:center; gap:10px; margin-top:18px; font-size:13px; }
  .pager button{ font-family:inherit; border:1px solid var(--border); background:var(--panel); border-radius:8px; padding:6px 14px; cursor:pointer; }
  .pager button:disabled{ opacity:.4; cursor:default; }

  footer{ text-align:center; color:var(--ink-soft); font-size:12px; padding:24px; }
</style>
</head>
<body>

<header>
  <div class="top">
    <div>
      <h1>لوحة تحكم منتجات متجر مكعب (Mokab)</h1>
      <p>سُحبت البيانات العامة الظاهرة للزوار من موقع mokab.com بتاريخ __GENERATED_AT__ — __TOTAL__ منتج</p>
    </div>
  </div>
</header>

<div class="wrap">

  <div class="stats-grid" id="statsGrid"></div>

  <div class="panel">
    <div class="analysis-cols">
      <div>
        <h2>أعلى التصنيفات من حيث عدد المنتجات</h2>
        <div id="topCategories"></div>
      </div>
      <div>
        <h2>أعلى العلامات التجارية من حيث عدد المنتجات</h2>
        <div id="topBrands"></div>
      </div>
    </div>
  </div>

  <div class="panel">
    <div class="controls">
      <input type="text" id="searchBox" placeholder="ابحث بالاسم أو SKU أو العلامة التجارية...">
      <select id="categoryFilter"><option value="">كل التصنيفات</option></select>
      <select id="availFilter">
        <option value="">كل حالات التوفر</option>
        <option value="in">متوفر فقط</option>
        <option value="out">غير متوفر فقط</option>
      </select>
      <select id="sortSelect">
        <option value="default">الترتيب الافتراضي</option>
        <option value="price_asc">السعر: من الأقل للأعلى</option>
        <option value="price_desc">السعر: من الأعلى للأقل</option>
        <option value="discount_desc">أعلى نسبة خصم</option>
        <option value="name_asc">الاسم أبجديًا</option>
      </select>
      <div class="view-toggle">
        <button id="btnCards" class="active">بطاقات</button>
        <button id="btnTable">جدول</button>
      </div>
    </div>
    <div class="result-count" id="resultCount"></div>
    <div id="cardsView" class="grid"></div>
    <div id="tableView" style="display:none; overflow-x:auto;">
      <table>
        <thead><tr>
          <th>صورة</th><th>الاسم</th><th>التصنيف</th><th>العلامة</th><th>SKU</th>
          <th>السعر الحالي</th><th>قبل الخصم</th><th>خصم%</th><th>التوفر</th><th>الضمان</th><th>رابط</th>
        </tr></thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
    <div class="pager">
      <button id="prevPage">السابق</button>
      <span id="pageInfo"></span>
      <button id="nextPage">التالي</button>
    </div>
  </div>

</div>

<footer>تم إنشاء هذه اللوحة تلقائيًا من بيانات mokab.com العامة — للاستخدام الداخلي فقط.</footer>

<script>
const PRODUCTS = __DATA_JSON__;
const PAGE_SIZE = 60;
let state = { view:'cards', page:1, filtered: PRODUCTS.slice() };

function fmt(n){ if(n===null||n===undefined) return ''; return Number(n).toLocaleString('ar-SA'); }

function computeStats(list){
  const total = list.length;
  const inStock = list.filter(p=>p.stock===true).length;
  const outStock = list.filter(p=>p.stock===false).length;
  const withPrice = list.filter(p=>p.price!=null);
  const avgPrice = withPrice.length ? (withPrice.reduce((a,p)=>a+p.price,0)/withPrice.length) : 0;
  const discounted = list.filter(p=>p.disc && p.disc>0);
  const avgDisc = discounted.length ? (discounted.reduce((a,p)=>a+p.disc,0)/discounted.length) : 0;
  const cats = new Set(list.map(p=>p.cat).filter(Boolean));
  const brands = new Set(list.map(p=>p.brand).filter(Boolean));
  const rated = list.filter(p=>p.rating!=null);
  const avgRating = rated.length ? (rated.reduce((a,p)=>a+p.rating,0)/rated.length) : null;

  document.getElementById('statsGrid').innerHTML = `
    <div class="stat-card"><div class="num">${fmt(total)}</div><div class="label">إجمالي المنتجات</div></div>
    <div class="stat-card"><div class="num">${fmt(inStock)}</div><div class="label">متوفر</div></div>
    <div class="stat-card"><div class="num">${fmt(outStock)}</div><div class="label">غير متوفر</div></div>
    <div class="stat-card"><div class="num">${cats.size}</div><div class="label">تصنيف</div></div>
    <div class="stat-card"><div class="num">${brands.size}</div><div class="label">علامة تجارية</div></div>
    <div class="stat-card"><div class="num">${fmt(Math.round(avgPrice))}</div><div class="label">متوسط السعر (ر.س)</div></div>
    <div class="stat-card"><div class="num">${discounted.length}</div><div class="label">منتج عليه خصم</div></div>
    <div class="stat-card"><div class="num">${avgDisc? avgDisc.toFixed(1)+'%':'-'}</div><div class="label">متوسط نسبة الخصم</div></div>
    ${avgRating!==null ? `<div class="stat-card"><div class="num">${avgRating.toFixed(2)}</div><div class="label">متوسط التقييم</div></div>` : ''}
  `;

  function topN(getKey, n){
    const counts = {};
    list.forEach(p=>{ const k=getKey(p); if(k){ counts[k]=(counts[k]||0)+1; } });
    return Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,n);
  }
  function renderBars(elId, entries){
    const max = entries.length ? entries[0][1] : 1;
    document.getElementById(elId).innerHTML = entries.map(([name,count])=>`
      <div class="bar-row">
        <div class="name" title="${name}">${name}</div>
        <div class="track"><div class="fill" style="width:${(count/max*100).toFixed(0)}%"></div></div>
        <div class="count">${count}</div>
      </div>`).join('') || '<div style="color:var(--ink-soft);font-size:13px;">لا توجد بيانات</div>';
  }
  renderBars('topCategories', topN(p=>p.cat, 8));
  renderBars('topBrands', topN(p=>p.brand, 8));
}

function populateCategoryFilter(){
  const cats = Array.from(new Set(PRODUCTS.map(p=>p.cat).filter(Boolean))).sort((a,b)=>a.localeCompare('ar'));
  const sel = document.getElementById('categoryFilter');
  cats.forEach(c=>{ const o=document.createElement('option'); o.value=c; o.textContent=c; sel.appendChild(o); });
}

function applyFilters(){
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const cat = document.getElementById('categoryFilter').value;
  const avail = document.getElementById('availFilter').value;
  const sortBy = document.getElementById('sortSelect').value;

  let list = PRODUCTS.filter(p=>{
    if(q){
      const hay = `${p.name||''} ${p.sku||''} ${p.brand||''}`.toLowerCase();
      if(!hay.includes(q)) return false;
    }
    if(cat && p.cat!==cat) return false;
    if(avail==='in' && p.stock!==true) return false;
    if(avail==='out' && p.stock!==false) return false;
    return true;
  });

  if(sortBy==='price_asc') list.sort((a,b)=>(a.price??1e12)-(b.price??1e12));
  else if(sortBy==='price_desc') list.sort((a,b)=>(b.price??-1)-(a.price??-1));
  else if(sortBy==='discount_desc') list.sort((a,b)=>(b.disc||0)-(a.disc||0));
  else if(sortBy==='name_asc') list.sort((a,b)=>(a.name||'').localeCompare(b.name||'', 'ar'));

  state.filtered = list;
  state.page = 1;
  computeStats(list);
  render();
}

function productCardHtml(p){
  const img = p.img ? `<img src="${p.img}" loading="lazy" alt="">` : `<div class="no-image">لا توجد صورة</div>`;
  let price = '';
  if(p.price!=null){
    price = `<span class="price-now">${fmt(p.price)} ر.س</span>`;
    if(p.old && p.old>p.price){
      price += `<span class="price-old">${fmt(p.old)} ر.س</span>`;
      if(p.disc) price += `<span class="badge-discount">خصم ${p.disc}%</span>`;
    }
  }
  const availClass = p.stock===true?'in-stock':(p.stock===false?'out-of-stock':'');
  return `<a class="card" href="${p.url||'#'}" target="_blank" rel="noopener">
    <div class="thumb">${img}</div>
    <div class="body">
      <div class="badge ${availClass}">${p.avail||'غير محدد'}</div>
      <h3>${p.name||''}</h3>
      <div>${price}</div>
      <div class="meta">${p.cat||''}</div>
      <div class="meta">${p.brand||''}</div>
      <div class="meta">SKU: ${p.sku||'-'}</div>
    </div>
  </a>`;
}

function productRowHtml(p){
  const availClass = p.stock===true?'in-stock':(p.stock===false?'out-of-stock':'');
  return `<tr>
    <td>${p.img?`<img class="t-thumb" src="${p.img}" loading="lazy">`:''}</td>
    <td class="t-name">${p.name||''}</td>
    <td>${p.cat||''}</td>
    <td>${p.brand||''}</td>
    <td>${p.sku||''}</td>
    <td>${p.price!=null?fmt(p.price)+' ر.س':''}</td>
    <td>${(p.old && p.old>p.price)?fmt(p.old)+' ر.س':''}</td>
    <td>${p.disc?p.disc+'%':''}</td>
    <td><span class="badge ${availClass}">${p.avail||'غير محدد'}</span></td>
    <td>${p.warranty||''}</td>
    <td><a class="t-link" href="${p.url||'#'}" target="_blank" rel="noopener">فتح ↗</a></td>
  </tr>`;
}

function render(){
  const list = state.filtered;
  const totalPages = Math.max(1, Math.ceil(list.length / PAGE_SIZE));
  if(state.page > totalPages) state.page = totalPages;
  const start = (state.page-1)*PAGE_SIZE;
  const pageItems = list.slice(start, start+PAGE_SIZE);

  document.getElementById('resultCount').textContent = `${list.length.toLocaleString('ar-SA')} نتيجة`;
  document.getElementById('pageInfo').textContent = `صفحة ${state.page} من ${totalPages}`;
  document.getElementById('prevPage').disabled = state.page<=1;
  document.getElementById('nextPage').disabled = state.page>=totalPages;

  if(state.view==='cards'){
    document.getElementById('cardsView').style.display='grid';
    document.getElementById('tableView').style.display='none';
    document.getElementById('cardsView').innerHTML = pageItems.map(productCardHtml).join('');
  } else {
    document.getElementById('cardsView').style.display='none';
    document.getElementById('tableView').style.display='block';
    document.getElementById('tableBody').innerHTML = pageItems.map(productRowHtml).join('');
  }
}

document.getElementById('searchBox').addEventListener('input', applyFilters);
document.getElementById('categoryFilter').addEventListener('change', applyFilters);
document.getElementById('availFilter').addEventListener('change', applyFilters);
document.getElementById('sortSelect').addEventListener('change', applyFilters);
document.getElementById('btnCards').addEventListener('click', ()=>{ state.view='cards'; document.getElementById('btnCards').classList.add('active'); document.getElementById('btnTable').classList.remove('active'); render(); });
document.getElementById('btnTable').addEventListener('click', ()=>{ state.view='table'; document.getElementById('btnTable').classList.add('active'); document.getElementById('btnCards').classList.remove('active'); render(); });
document.getElementById('prevPage').addEventListener('click', ()=>{ if(state.page>1){ state.page--; render(); } });
document.getElementById('nextPage').addEventListener('click', ()=>{ state.page++; render(); });

populateCategoryFilter();
applyFilters();
</script>
</body>
</html>"""


def build_dashboard_html(products, out_path: str):
    """يبني لوحة تحكم HTML واحدة ذاتية الاحتواء (بيانات مضمّنة داخل الملف نفسه
    فلا تحتاج خادم محلي ولا تواجه مشاكل CORS عند فتحها مباشرة بدبل كلك):
    تحليل عام (إحصاءات + أعلى التصنيفات/العلامات)، ثم قائمة كل المنتجات
    بعرض بطاقات (صورة أمام كل منتج) أو جدول قابل للفرز والفلترة والبحث."""
    import datetime

    def slim(p):
        return {
            "id": p.id, "name": p.name,
            "price": p.price_current, "old": p.price_before_discount, "disc": p.discount_percent,
            "cat": p.category, "brand": p.brand, "sku": p.sku, "opts": p.options,
            "avail": p.availability, "stock": p.in_stock, "qty": p.quantity,
            "rating": p.rating_value, "reviews": p.review_count, "warranty": p.warranty,
            "img": (p.image_urls.split(" | ")[0] if p.image_urls else ""),
            "url": p.url,
        }

    data_json = json.dumps([slim(p) for p in products], ensure_ascii=False)
    # حماية من كسر وسم السكربت لو وُجد النص "</script" داخل أي بيانات نصية
    data_json = data_json.replace("</script", "<\\/script")

    generated_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    html = (
        DASHBOARD_TEMPLATE
        .replace("__DATA_JSON__", data_json)
        .replace("__TOTAL__", str(len(products)))
        .replace("__GENERATED_AT__", generated_at)
    )

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)


def save_output(products, fmt: str, out_path: str):
    records = [asdict(p) for p in products]

    if fmt == "dashboard":
        build_dashboard_html(products, out_path)
        return

    if fmt == "html":
        build_html_catalog(products, out_path)
        return

    if fmt == "json":
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        return

    import pandas as pd

    df = pd.DataFrame(records)
    column_order = [
        "id", "name", "price_current", "price_before_discount", "discount_percent",
        "category", "brand", "sku", "options", "availability", "in_stock",
        "quantity", "rating_value", "review_count", "warranty",
        "image_urls", "description", "url", "errors",
    ]
    df = df[[c for c in column_order if c in df.columns]]
    df.rename(columns={
        "id": "معرف المنتج",
        "name": "اسم المنتج",
        "price_current": "السعر الحالي",
        "price_before_discount": "السعر قبل الخصم",
        "discount_percent": "نسبة الخصم %",
        "category": "التصنيف",
        "brand": "العلامة التجارية",
        "sku": "SKU / رقم الموديل",
        "options": "الخيارات (لون/مقاس)",
        "availability": "حالة التوفر",
        "in_stock": "متوفر؟ (True/False)",
        "quantity": "الكمية المتاحة",
        "rating_value": "متوسط التقييم",
        "review_count": "عدد المراجعات",
        "warranty": "الضمان",
        "image_urls": "روابط الصور",
        "description": "الوصف",
        "url": "رابط المنتج",
        "errors": "أخطاء (إن وجدت)",
    }, inplace=True)

    if fmt == "csv":
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:  # xlsx
        df.to_excel(out_path, index=False, engine="openpyxl")


def main():
    # ملاحظة: عند التشغيل من IDLE (بدون سطر أوامر) تُستخدم القيم الافتراضية
    # المحددة أعلى الملف في قسم "إعدادات سريعة". سطر الأوامر (--limit ...)
    # يبقى متاحًا لمن يفضّل استخدامه ويتجاوز الإعدادات الافتراضية عند تمريره.
    parser = argparse.ArgumentParser(description="سحب بيانات منتجات mokab.com (سلة)")
    parser.add_argument("--limit", type=int, default=DEFAULT_LIMIT, help="أقصى عدد منتجات (للتجربة)")
    parser.add_argument("--category", type=str, default=None,
                         help="غير مستخدم حاليًا للفلترة (واجهة سلة للتصنيفات غير موثوقة)؛ "
                              "محفوظ للتوافق المستقبلي")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="عدد الاتصالات المتزامنة")
    parser.add_argument("--format", choices=["dashboard", "xlsx", "csv", "json", "html"], default=DEFAULT_FORMAT)
    parser.add_argument("--output", type=str, default=None, help="مسار ملف الحفظ")
    parser.add_argument("--download-images", action="store_true", default=DEFAULT_DOWNLOAD_IMAGES,
                         help="تحميل ملفات الصور فعليًا (وليس فقط الروابط) إلى مجلد images/")
    parser.add_argument("--images-dir", type=str, default=None,
                         help="المجلد الذي تُحفظ فيه الصور عند استخدام --download-images")
    args, _unknown = parser.parse_known_args()

    file_ext = {"dashboard": "html", "html": "html", "xlsx": "xlsx", "csv": "csv", "json": "json"}[args.format]
    default_name = "mokab_dashboard.html" if args.format == "dashboard" else f"mokab_products.{file_ext}"
    out_path = args.output or os.path.join(SCRIPT_DIR, default_name)
    images_dir = (args.images_dir or os.path.join(SCRIPT_DIR, "images")) if args.download_images else None
    if images_dir:
        os.makedirs(images_dir, exist_ok=True)

    session = requests.Session()

    print("جاري جلب قائمة روابط كل المنتجات من sitemap...")
    product_list = get_all_product_urls(session, limit=args.limit)
    print(f"تم العثور على {len(product_list)} منتج. بدء السحب...")

    products = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {
            executor.submit(fetch_product, session, pid, url, images_dir): pid
            for pid, url in product_list
        }
        done = 0
        total = len(futures)
        for future in concurrent.futures.as_completed(futures):
            try:
                products.append(future.result())
            except Exception as e:
                pid = futures[future]
                products.append(Product(id=pid, errors=str(e)))
            done += 1
            if done % 25 == 0 or done == total:
                print(f"  {done}/{total} ...", file=sys.stderr)

    # إعادة الترتيب حسب الترتيب الأصلي في sitemap
    order = {pid: i for i, (pid, _) in enumerate(product_list)}
    products.sort(key=lambda p: order.get(p.id, 1_000_000))

    save_output(products, args.format, out_path)
    print(f"تم الحفظ في: {out_path}")

    error_count = sum(1 for p in products if p.errors)
    if error_count:
        print(f"تنبيه: {error_count} منتج واجه أخطاء جزئية أثناء السحب (راجع عمود 'أخطاء').")


if __name__ == "__main__":
    main()
