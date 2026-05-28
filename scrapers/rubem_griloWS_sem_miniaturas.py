# -*- coding: utf-8 -*-
"""Scraper simples por artista para o Oráculo Cultural.

Gera Excel no padrão histórico:
Título, Autor, Ano, Técnica, Dimensões, Preço, Descrição, Link da obra, Link da imagem da obra.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlsplit, urlunsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

AUTHOR_NAME = "Rubem Grilo"
START_URL = "https://www.arrematearte.com.br/artistas/rubem-grilo-1946"
SOURCE_SITE = "ArremateArte"
DEFAULT_TECHNIQUE = "Xilogravura"
OUTPUT_DEFAULT = "artes_Rubem_Grilo.xlsx"
SEM_INFO = "Sem informação"

COLUMNS = [
    "Título", "Autor", "Ano", "Técnica", "Dimensões", "Preço",
    "Descrição", "Link da obra", "Link da imagem da obra",
]

PRICE_RE = re.compile(r"R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}|R\$\s*\d+(?:[.,]\d{2})?", re.I)
YEAR_RE = re.compile(r"\b(18\d{2}|19\d{2}|20\d{2})\b")
DIM_RE = re.compile(
    r"\b\d{1,4}(?:[,.]\d{1,2})?\s*(?:x|×|X|por)\s*\d{1,4}(?:[,.]\d{1,2})?"
    r"(?:\s*(?:x|×|X|por)\s*\d{1,4}(?:[,.]\d{1,2})?)?\s*(?:cm|CM)?\b",
    re.I,
)
BAD_IMG_RE = re.compile(r"logo|avatar|sprite|icon|whatsapp|facebook|instagram|twitter|placeholder|blank|pagamento|selo|seguranca|segurança", re.I)
IMG_EXT_RE = re.compile(r"\.(?:jpe?g|png|webp|avif)(?:[?#].*)?$", re.I)

BLACKLIST_PATHS = (
    "/artistas", "/artista", "/blog", "/sobre", "/contato", "/checkout", "/cart", "/carrinho",
    "/minha-conta", "/politica", "/termos", "/leiloes", "/leilões", "/categorias", "/obras/pinturas",
)

@dataclass
class Obra:
    titulo: str = SEM_INFO
    autor: str = AUTHOR_NAME
    ano: str = SEM_INFO
    tecnica: str = DEFAULT_TECHNIQUE
    dimensoes: str = SEM_INFO
    preco: Any = SEM_INFO
    descricao: str = SEM_INFO
    link_obra: str = SEM_INFO
    link_imagem: str = SEM_INFO

    def as_row(self) -> dict[str, Any]:
        return {
            "Título": self.titulo,
            "Autor": self.autor,
            "Ano": self.ano,
            "Técnica": self.tecnica,
            "Dimensões": self.dimensoes,
            "Preço": self.preco,
            "Descrição": self.descricao,
            "Link da obra": self.link_obra,
            "Link da imagem da obra": self.link_imagem,
        }


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\ufeff", " ").replace("&nbsp;", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def sem_info(value: Any) -> str:
    text = clean(value)
    return text if text else SEM_INFO


def abs_url(url: str | None, base: str = START_URL) -> str:
    if not url:
        return ""
    url = str(url).strip()
    if url.startswith("//"):
        return "https:" + url
    return urljoin(base, url)


def strip_url(url: str) -> str:
    parts = urlsplit(abs_url(url))
    return urlunsplit((parts.scheme, parts.netloc, parts.path.rstrip("/"), "", "", ""))


def price_to_float(text: str) -> float | str:
    text = clean(text)
    m = PRICE_RE.search(text)
    if not m:
        return SEM_INFO
    raw = m.group(0).replace("R$", "").strip()
    raw = raw.replace(".", "").replace(",", ".")
    try:
        return float(raw)
    except Exception:
        return SEM_INFO


def extract_price(text: str) -> float | str:
    # Em páginas com promoção, preferimos o preço promocional.
    promo = re.search(r"Preço\s+Promocional\s*(R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}|R\$\s*\d+(?:[.,]\d{2})?)", text, flags=re.I)
    if promo:
        return price_to_float(promo.group(1))
    vals = PRICE_RE.findall(text)
    if vals:
        return price_to_float(vals[-1] if "Preço Promocional" in text else vals[0])
    return SEM_INFO


def format_dim_number(value: str) -> str:
    value = value.replace(".", ",")
    if "," in value:
        value = value.rstrip("0").rstrip(",")
    return value


def extract_dimensions(text: str) -> str:
    text = clean(text).replace("×", "x")
    m = DIM_RE.search(text)
    if not m:
        return SEM_INFO
    nums = re.findall(r"\d{1,4}(?:[,.]\d{1,2})?", m.group(0))[:3]
    nums = [format_dim_number(n) for n in nums]
    if len(nums) >= 3:
        return f"{nums[0]} x {nums[1]} x {nums[2]} cm"
    if len(nums) >= 2:
        return f"{nums[0]} x {nums[1]} cm"
    return SEM_INFO


def extract_year(text: str) -> str:
    # Evita pegar anos de biografia quando possível buscando perto de técnica/dimensão.
    for pat in [r"(?:Ano|Data|Date|Década de)\s*:?\s*(18\d{2}|19\d{2}|20\d{2})", r"(18\d{2}|19\d{2}|20\d{2})\s*,?\s*\d{1,4}\s*x"]:
        m = re.search(pat, text, flags=re.I)
        if m:
            return m.group(1)
    m = YEAR_RE.search(text)
    return m.group(1) if m else SEM_INFO


def normalize_technique(text: str) -> str:
    t = clean(text).lower()
    repl = str.maketrans("áàãâéêíóôõúç", "aaaaeeiooouc")
    tn = t.translate(repl)
    if "xilograv" in tn or "woodcut" in tn:
        return "Xilogravura"
    if "linograv" in tn or "linocut" in tn:
        return "Linogravura"
    if "agua-forte" in tn or "agua forte" in tn or "etching" in tn:
        return "Água-forte"
    if "agua-tinta" in tn or "aguatinta" in tn or "aquatint" in tn:
        return "Água-tinta"
    if "ponta-seca" in tn or "ponta seca" in tn or "drypoint" in tn:
        return "Ponta-seca"
    if "gravura em metal" in tn or "metal" in tn and "grav" in tn:
        return "Gravura em metal"
    if "litograv" in tn or "litograf" in tn:
        return "Litografia"
    if "serigraf" in tn:
        return "Serigrafia"
    if "oleo" in tn or "oil" in tn:
        return "Pintura a óleo"
    if "acril" in tn:
        return "Pintura acrílica"
    if "aquarela" in tn:
        return "Aquarela"
    if "guache" in tn:
        return "Guache"
    if "gravura" in tn or "gravad" in tn:
        return "Gravura"
    if "desenho" in tn:
        return "Desenho"
    if "pintura" in tn:
        return "Pintura"
    return DEFAULT_TECHNIQUE if DEFAULT_TECHNIQUE != SEM_INFO else "Outros"


def make_session() -> requests.Session:
    retry = Retry(total=3, connect=3, read=3, status=3, backoff_factor=0.6, status_forcelist=(403, 429, 500, 502, 503, 504), allowed_methods=("GET", "HEAD"), raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retry)
    s = requests.Session()
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control": "no-cache",
    })
    return s


def fetch(session: requests.Session, url: str, timeout: int = 35) -> str:
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    if not resp.encoding or resp.encoding.lower() in {"iso-8859-1", "ascii"}:
        resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def render_with_playwright(url: str, timeout_ms: int = 70000) -> str:
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        return ""
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            page = browser.new_page(viewport={"width": 1440, "height": 1200}, user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/125 Safari/537.36")
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            try:
                page.wait_for_load_state("networkidle", timeout=12000)
            except Exception:
                pass
            for _ in range(6):
                page.mouse.wheel(0, 1800)
                page.wait_for_timeout(800)
            html = page.content()
            browser.close()
            return html
    except Exception as exc:
        print(f"Playwright fallback falhou em {url}: {exc}", file=sys.stderr)
        return ""


def soup_lines(soup: BeautifulSoup) -> list[str]:
    return [clean(x) for x in soup.get_text("\n", strip=True).splitlines() if clean(x)]


def parse_json_scripts(soup: BeautifulSoup) -> list[Any]:
    out = []
    for sc in soup.select("script[type='application/ld+json'], script[type='application/json']"):
        raw = clean(sc.string or sc.get_text(" ", strip=True))
        if not raw:
            continue
        try:
            out.append(json.loads(raw))
        except Exception:
            pass
    return out


def flatten(obj: Any):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from flatten(v)
    elif isinstance(obj, list):
        for item in obj:
            yield from flatten(item)


def json_first(jsons: list[Any], keys: tuple[str, ...]) -> str:
    for obj in flatten(jsons):
        if not isinstance(obj, dict):
            continue
        for k in keys:
            v = obj.get(k)
            if isinstance(v, str) and clean(v):
                return clean(v)
            if isinstance(v, dict):
                for kk in ("name", "url", "contentUrl"):
                    if isinstance(v.get(kk), str) and clean(v.get(kk)):
                        return clean(v.get(kk))
            if isinstance(v, list):
                for item in v:
                    if isinstance(item, str) and clean(item):
                        return clean(item)
                    if isinstance(item, dict):
                        for kk in ("name", "url", "contentUrl"):
                            if isinstance(item.get(kk), str) and clean(item.get(kk)):
                                return clean(item.get(kk))
    return ""


def extract_image(soup: BeautifulSoup, base_url: str, title: str = "") -> str:
    candidates: list[tuple[int, str]] = []
    def add(url: str, score: int):
        url = abs_url(url, base_url)
        if not url or url.startswith("data:"):
            return
        low = url.lower()
        if BAD_IMG_RE.search(low):
            score -= 60
        if IMG_EXT_RE.search(low):
            score += 25
        candidates.append((score, url))

    for sel in ["meta[property='og:image']", "meta[property='og:image:secure_url']", "meta[name='twitter:image']", "link[rel='image_src']"]:
        node = soup.select_one(sel)
        if node:
            add(node.get("content") or node.get("href") or "", 120)

    title_words = {w for w in re.split(r"\W+", title.lower()) if len(w) >= 4}
    for img in soup.select("img"):
        score = 50
        alt = clean(img.get("alt") or img.get("title") or "")
        if title_words and title_words.intersection(re.split(r"\W+", alt.lower())):
            score += 35
        for attr in ["src", "data-src", "data-original", "data-lazy-src", "data-full", "data-zoom-image"]:
            if img.get(attr):
                add(img.get(attr), score)
        srcset = img.get("srcset") or img.get("data-srcset")
        if srcset:
            parts = [p.strip().split()[0] for p in srcset.split(",") if p.strip()]
            if parts:
                add(parts[-1], score + 10)
    if not candidates:
        return SEM_INFO
    best = sorted(candidates, key=lambda x: x[0], reverse=True)[0][1]
    return best or SEM_INFO


def title_from_meta(soup: BeautifulSoup) -> str:
    for sel in ["meta[property='og:title']", "meta[name='twitter:title']"]:
        n = soup.select_one(sel)
        if n and clean(n.get("content")):
            return clean(n.get("content"))
    if soup.select_one("h1"):
        return clean(soup.select_one("h1").get_text(" ", strip=True))
    if soup.title and soup.title.string:
        return clean(soup.title.string)
    return ""


def clean_title(raw: str, author: str = AUTHOR_NAME) -> str:
    title = clean(raw)
    title = re.sub(r"\s*[-|].*$", "", title) if SOURCE_SITE == "ArteGaleria" and "Milan" not in title else title
    title = re.sub(r"\s+obra\s+de\s+.+$", "", title, flags=re.I)
    title = re.sub(r"\s*[-|]\s*(Blomb[oô]|Arremate Arte|Oto Reifschneider|Galeria de Arte).*$", "", title, flags=re.I)
    title = re.sub(rf"\s*[|l]\s*{re.escape(author)}\s*$", "", title, flags=re.I)
    title = title.strip(" -–—|,;:")
    return title if title else SEM_INFO


def parse_blombo_detail(soup: BeautifulSoup, url: str, title_hint: str = "", price_hint: Any = SEM_INFO, img_hint: str = "") -> Obra:
    text = clean(soup.get_text(" ", strip=True))
    meta_title = title_from_meta(soup)
    title = clean_title(meta_title) or clean_title(title_hint)
    if title == AUTHOR_NAME or title == SEM_INFO:
        # linha de breadcrumb antes do bloco técnico
        m = re.search(r"Obras\s+(?:Pinturas|Gravuras e múltiplos|Desenhos)?\s*([^#]{2,120}?)\s*(?:Cícero Dias|Cicero Dias)", text, re.I)
        if m:
            title = clean_title(m.group(1))
    # Produto Blombo: "Título  Técnica  dimensões".
    dim = extract_dimensions(text)
    tecnica_src = text
    if title != SEM_INFO:
        idx = text.lower().find(title.lower())
        if idx >= 0:
            tecnica_src = text[idx: idx + 240]
    tecnica = normalize_technique(tecnica_src)
    ano = extract_year(tecnica_src)
    price = extract_price(text)
    if price == SEM_INFO:
        price = price_hint
    image = extract_image(soup, url, title)
    if image == SEM_INFO:
        image = img_hint or SEM_INFO
    desc = SEM_INFO
    # Blombo costuma repetir a biografia depois da obra; manter descrição curta evita poluição.
    return Obra(title, AUTHOR_NAME, ano, tecnica, dim, price, desc, url, image)


def parse_generic_detail(soup: BeautifulSoup, url: str, title_hint: str = "", price_hint: Any = SEM_INFO, img_hint: str = "") -> Obra:
    text = clean(soup.get_text(" ", strip=True))
    lines = soup_lines(soup)
    title = clean_title(title_from_meta(soup))
    if title in {SEM_INFO, AUTHOR_NAME}:
        title = clean_title(title_hint)
    # Preço, dimensão, técnica e ano.
    price = extract_price(text)
    if price == SEM_INFO:
        price = price_hint
    dim = extract_dimensions(text)
    tecnica = normalize_technique(" ".join(lines[:80]) + " " + title)
    ano = extract_year(text)
    img = extract_image(soup, url, title)
    if img == SEM_INFO:
        img = img_hint or SEM_INFO
    desc = SEM_INFO
    # Tenta uma descrição curta só quando houver ficha técnica clara.
    for node in soup.select(".woocommerce-product-details__short-description, .summary, .product-info__description-inner, [class*='description']"):
        d = clean(node.get_text(" ", strip=True))
        if len(d) > 20 and len(d) < 600:
            desc = d
            break
    return Obra(title, AUTHOR_NAME, ano, tecnica, dim, price, desc, url, img)


def is_probable_product_link(url: str, text: str) -> bool:
    u = strip_url(url)
    parts = urlsplit(u)
    path = parts.path.lower()
    if not path or path == "/":
        return False
    if any(path.startswith(p) for p in BLACKLIST_PATHS):
        return False
    if SOURCE_SITE == "ArteGaleria":
        return "/produto/" in path
    if SOURCE_SITE == "Blombo":
        if path.startswith("/artistas") or path.startswith("/obras") or path.startswith("/blog"):
            return False
        return bool(re.search(r"-\d{3,}$", path)) or AUTHOR_NAME.lower().split()[0] in clean(text).lower()
    if SOURCE_SITE == "ArremateArte":
        if any(x in path for x in ["/lotes", "/lote", "/peca", "/pecas", "/item", "/obra", "/produto"]):
            return True
        # Alguns lotes usam o slug do artista no próprio href.
        return "rubem-grilo" in path and not path.endswith("/artistas/rubem-grilo-1946")
    return False


def extract_listing_items(soup: BeautifulSoup, page_url: str) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    seen: set[str] = set()
    for a in soup.select("a[href]"):
        href = abs_url(a.get("href"), page_url)
        link = strip_url(href)
        label = clean(a.get_text(" ", strip=True) or a.get("title") or a.get("aria-label") or "")
        if not is_probable_product_link(link, label):
            continue
        if link in seen:
            continue
        seen.add(link)
        parent: Tag | BeautifulSoup | None = a
        context = label
        for _ in range(5):
            if parent is None:
                break
            txt = clean(parent.get_text(" ", strip=True)) if isinstance(parent, Tag) else ""
            if len(txt) > len(context):
                context = txt
            # Para cards de produto, essa faixa costuma conter título + preço + técnica.
            if PRICE_RE.search(txt) or len(txt) > 80:
                break
            parent = parent.parent
        title = clean_title(label)
        price = extract_price(context)
        dim = extract_dimensions(context)
        tecnica = normalize_technique(context + " " + title)
        img = SEM_INFO
        if isinstance(parent, Tag):
            img = extract_image(BeautifulSoup(str(parent), "lxml"), page_url, title)
        items.append({"link": link, "title": title, "price": price, "dimensions": dim, "technique": tecnica, "image": img, "context": context})
    return items


def rows_from_listing_text(soup: BeautifulSoup, page_url: str) -> list[Obra]:
    # Fallback para páginas em que o web/html já traz linhas completas, como Blombo.
    text = soup.get_text("\n", strip=True)
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    rows: list[Obra] = []
    if SOURCE_SITE != "Blombo":
        return rows
    for i, line in enumerate(lines):
        if AUTHOR_NAME.lower() != line.lower():
            continue
        chunk = " ".join(lines[i:i+8])
        dim = extract_dimensions(chunk)
        if dim == SEM_INFO:
            continue
        # linha após autor normalmente é o título, às vezes a mesma ancora aparece antes.
        title = SEM_INFO
        for cand in lines[i+1:i+4]:
            if cand and not PRICE_RE.search(cand) and "+ DETALHES" not in cand and extract_dimensions(cand) == SEM_INFO:
                title = clean_title(cand)
                break
        price = extract_price(chunk)
        tech = normalize_technique(chunk)
        rows.append(Obra(title, AUTHOR_NAME, extract_year(chunk), tech, dim, price, SEM_INFO, page_url, SEM_INFO))
    return rows


def detail_obras(session: requests.Session, listing_items: list[dict[str, Any]], max_obras: int, delay: float) -> list[Obra]:
    obras: list[Obra] = []
    if max_obras and max_obras > 0:
        listing_items = listing_items[:max_obras]
    for idx, item in enumerate(listing_items, start=1):
        link = item["link"]
        print(f"[{idx}/{len(listing_items)}] Detalhando: {link}")
        try:
            html = fetch(session, link)
        except Exception as exc:
            print(f"  Falha requests; tentando renderizar: {exc}")
            html = render_with_playwright(link)
        if not html:
            obras.append(Obra(item.get("title") or SEM_INFO, AUTHOR_NAME, SEM_INFO, item.get("technique") or DEFAULT_TECHNIQUE, item.get("dimensions") or SEM_INFO, item.get("price") or SEM_INFO, SEM_INFO, link, item.get("image") or SEM_INFO))
            continue
        soup = BeautifulSoup(html, "lxml")
        if SOURCE_SITE == "Blombo":
            obra = parse_blombo_detail(soup, link, item.get("title") or "", item.get("price"), item.get("image") or "")
        else:
            obra = parse_generic_detail(soup, link, item.get("title") or "", item.get("price"), item.get("image") or "")
        # Completa com dados da listagem quando o detalhe não traz tudo.
        if obra.dimensoes == SEM_INFO and item.get("dimensions"):
            obra.dimensoes = item["dimensions"]
        if obra.tecnica in {SEM_INFO, "Outros"} and item.get("technique"):
            obra.tecnica = item["technique"]
        if obra.preco == SEM_INFO and item.get("price"):
            obra.preco = item["price"]
        if obra.link_imagem == SEM_INFO and item.get("image"):
            obra.link_imagem = item["image"]
        obras.append(obra)
        if delay:
            time.sleep(delay)
    return obras


def scrape(start_url: str, max_obras: int = 0, max_pages: int = 0, delay: float = 0.3) -> pd.DataFrame:
    session = make_session()
    try:
        html = fetch(session, start_url)
    except Exception as exc:
        print(f"Falha requests na listagem; tentando Playwright: {exc}")
        html = render_with_playwright(start_url)
    if not html:
        return pd.DataFrame(columns=COLUMNS)
    soup = BeautifulSoup(html, "lxml")
    items = extract_listing_items(soup, start_url)
    if not items:
        # Fallback renderizado para sites SPA.
        rendered = render_with_playwright(start_url)
        if rendered:
            soup = BeautifulSoup(rendered, "lxml")
            items = extract_listing_items(soup, start_url)
    obras = detail_obras(session, items, max_obras, delay) if items else []
    if not obras:
        obras = rows_from_listing_text(soup, start_url)
        if max_obras and max_obras > 0:
            obras = obras[:max_obras]
    df = pd.DataFrame([o.as_row() for o in obras], columns=COLUMNS)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link da obra", "Título"], keep="first")
    return df


def exportar(df: pd.DataFrame, output: str) -> None:
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = SEM_INFO
    df = df.reindex(columns=COLUMNS)
    ext = os.path.splitext(output)[1].lower()
    if ext in {".xlsx", ".xlsm", ".xls"}:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Dados")
            ws = writer.sheets["Dados"]
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 70)
            ws.freeze_panes = "A2"
        return
    df.to_csv(output, index=False, encoding="utf-8-sig")


def main() -> None:
    parser = argparse.ArgumentParser(description=f"Scraper de obras de {AUTHOR_NAME}")
    parser.add_argument("--url", default=os.getenv("AUTOR_URL", START_URL), help="URL da página do artista")
    parser.add_argument("--output", default=os.getenv("OUTPUT", OUTPUT_DEFAULT), help="Arquivo de saída .xlsx ou .csv")
    parser.add_argument("--max-obras", type=int, default=int(os.getenv("MAX_OBRAS", "0")), help="Limite de obras. 0 = sem limite")
    parser.add_argument("--max-pages", type=int, default=int(os.getenv("MAX_PAGES", "0")), help="Compatibilidade; não usado na coleta simples")
    parser.add_argument("--delay", type=float, default=float(os.getenv("DELAY", "0.3")), help="Pausa entre requisições")
    args = parser.parse_args()

    started = datetime.now()
    df = scrape(args.url, max_obras=args.max_obras, max_pages=args.max_pages, delay=args.delay)
    exportar(df, args.output)
    print(f"Concluído: {len(df)} obras de {AUTHOR_NAME} exportadas para {args.output} em {datetime.now() - started}.")


if __name__ == "__main__":
    main()
