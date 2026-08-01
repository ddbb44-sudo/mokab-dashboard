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
