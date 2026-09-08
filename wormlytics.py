#!/usr/bin/env python3
"""
██╗    ██╗ ██████╗ ██████╗ ███╗   ███╗██╗  ██╗   ██╗████████╗██╗ ██████╗███████╗
██║    ██║██╔═══██╗██╔══██╗████╗ ████║██║  ╚██╗ ██╔╝╚══██╔══╝██║██╔════╝██╔════╝
██║ █╗ ██║██║   ██║██████╔╝██╔████╔██║██║   ╚████╔╝    ██║   ██║██║     ███████╗
██║███╗██║██║   ██║██╔══██╗██║╚██╔╝██║██║    ╚██╔╝     ██║   ██║██║     ╚════██║
╚███╔███╔╝╚██████╔╝██║  ██║██║ ╚═╝ ██║███████╗██║      ██║   ██║╚██████╗███████║
 ╚══╝╚══╝  ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝╚═╝      ╚═╝   ╚═╝ ╚═════╝╚══════╝

Analytics Data Layer Hunter — v5.5
Human-First CDP Mode + Autonomous Navigation + Excel Output + E-commerce Parser
"""

import asyncio
import json
import os
import platform
import random
import re
import subprocess
import sys
import time
import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse, urljoin

try:
    from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.text import Text
    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
    from rich import box
    import rich.traceback
    rich.traceback.install()
except ImportError as e:
    print(f"[ERROR] {e}\nEjecuta: pip install playwright rich openpyxl && playwright install chromium")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                                  GradientFill)
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table as XlTable, TableStyleInfo
    XLSX_OK = True
except ImportError:
    XLSX_OK = False

console = Console()

CDP_PORT     = 9222
NAV_DURATION = 300   # 5 minutos de navegación autónoma

BANNER = """[bold green]
██╗    ██╗ ██████╗ ██████╗ ███╗   ███╗██╗  ██╗   ██╗████████╗██╗ ██████╗███████╗
██║    ██║██╔═══██╗██╔══██╗████╗ ████║██║  ╚██╗ ██╔╝╚══██╔══╝██║██╔════╝██╔════╝
██║ █╗ ██║██║   ██║██████╔╝██╔████╔██║██║   ╚████╔╝    ██║   ██║██║     ███████╗
██║███╗██║██║   ██║██╔══██╗██║╚██╔╝██║██║    ╚██╔╝     ██║   ██║██║     ╚════██║
╚███╔███╔╝╚██████╔╝██║  ██║██║ ╚═╝ ██║███████╗██║      ██║   ██║╚██████╗███████║
 ╚══╝╚══╝  ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝╚═╝      ╚═╝   ╚═╝ ╚═════╝╚══════╝
[/bold green][dim]Analytics Data Layer Hunter — v5.5  |  CDP + Autonomous Nav + Excel + E-commerce[/dim]"""

# ─────────────────────────────────────────────────────────────────────────────
# CHROME DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def get_chrome_path():
    s = platform.system()
    paths = {
        "Windows": [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        ],
        "Darwin": [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ],
        "Linux": [
            "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium", "/usr/bin/chromium-browser", "/snap/bin/chromium",
        ],
    }
    for p in paths.get(s, []):
        if os.path.exists(p):
            return p
    return None

# ─────────────────────────────────────────────────────────────────────────────
# JS SCRIPTS
# ─────────────────────────────────────────────────────────────────────────────

JS_INJECT = """
() => {
    if (window.__wl_injected) return 'already';
    window.__wl_injected = true;
    window.__wl_events   = [];
    window.__wl_beacons  = [];

    // GTM dataLayer.push
    const patchDL = () => {
        if (typeof dataLayer !== 'undefined' && !dataLayer.__wl_patched) {
            const orig = dataLayer.push.bind(dataLayer);
            dataLayer.push = function(...args) {
                args.forEach(evt => {
                    try { window.__wl_events.push({ source:'GTM_dataLayer', ts: new Date().toISOString(), data: JSON.parse(JSON.stringify(evt)) }); } catch(e){}
                });
                return orig(...args);
            };
            dataLayer.__wl_patched = true;
        }
    };
    patchDL();
    setTimeout(patchDL, 1500);
    setTimeout(patchDL, 3500);

    // Tealium utag.track / utag.view / utag.link
    const patchTealium = () => {
        if (typeof utag !== 'undefined' && !utag.__wl_patched) {
            ['track','view','link'].forEach(m => {
                if (typeof utag[m] === 'function') {
                    const orig = utag[m].bind(utag);
                    utag[m] = function(eventType, dataObj, ...rest) {
                        try {
                            window.__wl_events.push({
                                source: `Tealium_utag.${m}`,
                                ts: new Date().toISOString(),
                                data: { call: m, eventType, payload: JSON.parse(JSON.stringify(dataObj || {})) }
                            });
                        } catch(e){}
                        return orig(eventType, dataObj, ...rest);
                    };
                }
            });
            utag.__wl_patched = true;
        }
    };
    setTimeout(patchTealium, 500);
    setTimeout(patchTealium, 2000);
    setTimeout(patchTealium, 4000);

    // Segment analytics.track / page / identify
    const patchSegment = () => {
        if (typeof analytics !== 'undefined' && analytics.track && !analytics.__wl_patched) {
            ['track','page','identify','group'].forEach(m => {
                if (typeof analytics[m] === 'function') {
                    const orig = analytics[m].bind(analytics);
                    analytics[m] = function(...args) {
                        try {
                            window.__wl_events.push({
                                source: `Segment_analytics.${m}`,
                                ts: new Date().toISOString(),
                                data: { call: m, args: args.slice(0,3).map(a => {
                                    try{ return JSON.parse(JSON.stringify(a)); }catch(e){ return String(a); }
                                })}
                            });
                        } catch(e){}
                        return orig(...args);
                    };
                }
            });
            analytics.__wl_patched = true;
        }
    };
    setTimeout(patchSegment, 500);
    setTimeout(patchSegment, 2500);

    // Piano Analytics / AT Internet
    const patchPiano = () => {
        if (typeof pa !== 'undefined' && pa.sendEvent && !pa.__wl_patched) {
            const orig = pa.sendEvent.bind(pa);
            pa.sendEvent = function(eventName, props, ...rest) {
                try {
                    window.__wl_events.push({
                        source: 'Piano_pa.sendEvent',
                        ts: new Date().toISOString(),
                        data: { event: eventName, props: JSON.parse(JSON.stringify(props || {})) }
                    });
                } catch(e){}
                return orig(eventName, props, ...rest);
            };
            pa.__wl_patched = true;
        }
    };
    setTimeout(patchPiano, 1000);
    setTimeout(patchPiano, 3000);

    // Adobe Web SDK (alloy) — intercept alloy() calls and sendEvent
    const patchAlloy = () => {
        if (typeof alloy !== 'undefined' && !window.__wl_alloy_patched) {
            const origAlloy = window.alloy;
            window.alloy = function(command, payload) {
                try {
                    if (command === 'sendEvent' || command === 'collectEvent') {
                        const snap = {};
                        // Extract __adobe.analytics vars from XDM payload
                        const aa = (payload?.data?.__adobe?.analytics) || {};
                        Object.assign(snap, aa);
                        // Extract xdm fields useful for analytics
                        const xdm = payload?.xdm || {};
                        if (xdm.eventType) snap['xdm.eventType'] = xdm.eventType;
                        if (xdm.web?.webPageDetails?.name) snap['pageName'] = xdm.web.webPageDetails.name;
                        if (xdm.web?.webPageDetails?.URL)  snap['pageURL']  = xdm.web.webPageDetails.URL;
                        if (xdm.commerce) snap['xdm.commerce'] = JSON.stringify(xdm.commerce).slice(0,200);
                        if (!window.__wl_s_snapshots) window.__wl_s_snapshots = [];
                        window.__wl_s_snapshots.push(snap);
                        window.__wl_events.push({
                            source: 'Alloy_sendEvent',
                            ts: new Date().toISOString(),
                            data: { command, xdm_eventType: xdm.eventType || null,
                                    pageName: xdm.web?.webPageDetails?.name || null,
                                    aa_vars: snap }
                        });
                    }
                } catch(e){}
                return origAlloy(command, payload);
            };
            // Copy all properties from original alloy to the wrapper
            try { Object.assign(window.alloy, origAlloy); } catch(e){}
            window.__wl_alloy_patched = true;
        }
    };
    setTimeout(patchAlloy, 300);
    setTimeout(patchAlloy, 1500);
    setTimeout(patchAlloy, 3500);

    // Adobe s.t / s.tl — snapshot completo del s object en el momento del call
    const snapS = () => {
        const so = window.s;
        const snap = {};
        try {
            for (let i=1;i<=250;i++){const k=`eVar${i}`;if(so[k]!==undefined&&so[k]!==''&&so[k]!==null)snap[k]=so[k];}
            for (let i=1;i<=75; i++){const k=`prop${i}`;if(so[k]!==undefined&&so[k]!==''&&so[k]!==null)snap[k]=so[k];}
            ['pageName','pageURL','channel','server','hier1','hier2','hier3','campaign','products',
             'purchaseID','visitorID','charSet','pageType','transactionID','zip','state','country',
             'linkName','linkType','currencyCode','events'].forEach(k=>{if(so[k])snap[k]=so[k];});
        } catch(e){}
        return snap;
    };
    const patchS = () => {
        if (typeof s !== 'undefined' && !s.__wl_patched) {
            ['t','tl'].forEach(m => {
                if (typeof s[m] === 'function') {
                    const orig = s[m].bind(s);
                    s[m] = function(...args) {
                        const snapshot = snapS();
                        window.__wl_events.push({
                            source: `AA_s.${m}`,
                            ts: new Date().toISOString(),
                            data: { method: m, s_snapshot: snapshot }
                        });
                        // También actualizar el store global con el snapshot más completo
                        if (!window.__wl_s_snapshots) window.__wl_s_snapshots = [];
                        window.__wl_s_snapshots.push(snapshot);
                        return orig(...args);
                    };
                }
            });
            s.__wl_patched = true;
        }
    };
    patchS();
    setTimeout(patchS, 1000);
    setTimeout(patchS, 2500);
    setTimeout(patchS, 4000);

    // Beacons via fetch / XHR
    const analyticsRe = /omtrdc|adobedc|google-analytics|googletagmanager|analytics[.]google/i;
    const origFetch = window.fetch;
    window.fetch = function(input, init) {
        const url = typeof input === 'string' ? input : (input?.url || '');
        if (analyticsRe.test(url)) window.__wl_beacons.push({ type:'fetch', url: url.slice(0,250), ts: new Date().toISOString() });
        return origFetch.apply(this, arguments);
    };
    const origOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function(method, url) {
        if (analyticsRe.test(String(url))) window.__wl_beacons.push({ type:'xhr', url: String(url).slice(0,250), ts: new Date().toISOString() });
        return origOpen.apply(this, arguments);
    };

    return 'ok';
}
"""

JS_EXTRACT = """
() => {
    const r = { adobe:{}, google:{}, gtm_containers:[], url: location.href, title: document.title };

    // s object — leer estado actual + merging con snapshots de s.t() calls
    const buildSObj = (so) => {
        const av = {};
        for (let i=1;i<=250;i++){const k=`eVar${i}`;if(so[k]!==undefined&&so[k]!==''&&so[k]!==null)av[k]=so[k];}
        for (let i=1;i<=75; i++){const k=`prop${i}`;if(so[k]!==undefined&&so[k]!==''&&so[k]!==null)av[k]=so[k];}
        ['pageName','pageURL','channel','server','hier1','hier2','hier3','campaign','products',
         'purchaseID','visitorID','charSet','pageType','transactionID','zip','state','country',
         'linkName','linkType','currencyCode','trackDownloadLinks','cookieDomainPeriods','events'].forEach(k=>{if(so[k])av[k]=so[k];});
        return av;
    };
    if (typeof s !== 'undefined') {
        let av = buildSObj(window.s);
        // Merge snapshots captured at s.t() time (más fiables que el estado actual)
        if (window.__wl_s_snapshots && window.__wl_s_snapshots.length > 0) {
            // El snapshot más reciente suele tener más variables populadas
            const bestSnap = window.__wl_s_snapshots.reduce((best, snap) =>
                Object.keys(snap).length > Object.keys(best).length ? snap : best,
                window.__wl_s_snapshots[0]
            );
            av = Object.assign({}, av, bestSnap);  // merge: snapshot wins for populated vars
        }
        r.adobe.s_object = av;
        r.adobe.report_suite = window.s.account || window.s.s_account || null;
        r.adobe.s_snapshots_count = (window.__wl_s_snapshots || []).length;
    }

    // _satellite
    if (typeof _satellite !== 'undefined') {
        const sat = { detected:true, property:null, environment:null, build_date:null, dataElements:{} };
        try {
            sat.property    = _satellite.property?.name || null;
            sat.environment = _satellite.environment?.stage || null;
            sat.build_date  = _satellite.buildInfo?.buildDate || null;
            // Extraer todos los DE del contenedor
            const deKeys = Object.keys(_satellite._container?.dataElements || {});
            deKeys.slice(0,100).forEach(de => {
                try { const v=_satellite.getVar(de); if(v!==undefined&&v!==null&&v!=='') sat.dataElements[de]=v; } catch(e){}
            });
        } catch(e) {}
        r.adobe.satellite = sat;
    }

    // alloy (Web SDK)
    if (typeof alloy !== 'undefined')
        r.adobe.web_sdk = { detected:true, version: window.__alloy_version || 'unknown' };

    // digitalData
    if (typeof digitalData !== 'undefined')
        try { r.adobe.digitalData = JSON.parse(JSON.stringify(digitalData)); } catch(e){}

    // ACDL
    if (typeof adobeDataLayer !== 'undefined')
        try { r.adobe.acdl = { detected:true, events: adobeDataLayer.slice(-20) }; } catch(e){}

    // GTM / GA
    if (typeof dataLayer !== 'undefined' && Array.isArray(dataLayer))
        r.google.dataLayer = dataLayer.slice(-50);
    if (typeof gtag !== 'undefined') r.google.gtag = { detected:true };
    if (window.google_tag_manager) r.gtm_containers = Object.keys(window.google_tag_manager);
    if (typeof ga !== 'undefined') {
        try {
            const tr = ga.getAll ? ga.getAll() : [];
            r.google.ua_trackers = tr.map(t=>({ name:t.get('name'), trackingId:t.get('trackingId'), clientId:t.get('clientId') }));
        } catch(e){ r.google.ua_detected=true; }
    }

    r.runtime_events   = window.__wl_events    || [];
    r.runtime_beacons  = window.__wl_beacons   || [];
    r.adobe.all_s_snapshots = window.__wl_s_snapshots || [];
    return r;
}
"""

JS_GET_LINKS = """
(domain) => {
    const skipWords = ['secure','login','signin','sign-in','logout','my-account',
                       'checkout','basket','payment','confirm','register','signup',
                       'sign-up','/auth/','?token=','?session='];
    const skipExts  = /[.](pdf|zip|jpg|jpeg|png|gif|svg|mp4|mp3)$/i;
    const links = Array.from(document.querySelectorAll('a[href]'))
        .map(a => { try { return new URL(a.href).href; } catch(e) { return null; } })
        .filter(href => href && href.startsWith('http'))
        .filter(href => { try { return new URL(href).hostname.includes(domain); } catch(e){ return false; } })
        .filter(href => !skipWords.some(w => href.includes(w)))
        .filter(href => !skipExts.test(href))
        .filter(href => href.indexOf('#') === -1);
    return [...new Set(links)].slice(0, 50);
}
"""

ANALYTICS_PATTERNS = {
    # Adobe
    "Adobe Analytics (AppMeasurement)": ["[.]2o7[.]net", "omtrdc[.]net", "b/ss/"],
    "Adobe Web SDK (alloy)":            ["adobedc[.]net", "edge[.]adobedc[.]net", "/ee/v2/", "/ee/or2/"],
    # Google
    "Google Analytics 4":               ["google-analytics[.]com/g/collect", "analytics[.]google[.]com/g/collect"],
    "Universal Analytics":              ["google-analytics[.]com/collect"],
    "Google Tag Manager":               ["googletagmanager[.]com/gtm[.]js", "googletagmanager[.]com/gtag/js"],
    # Tealium
    "Tealium iQ (TMS)":                 ["tags[.]tiqcdn[.]com", "tealiumiq[.]com", "tiq-cdn[.]com"],
    "Tealium AudienceStream":           ["visitor[.]tealiumiq[.]com", "collect[.]tealiumiq[.]com"],
    # Segment
    "Segment":                          ["cdn[.]segment[.]com", "api[.]segment[.]io", "analytics[.]js"],
    # Ensighten
    "Ensighten":                        ["nexus[.]ensighten[.]com", "[.]bootstrapper[.]js"],
    # Commanders Act (TagCommander)
    "Commanders Act":                   ["tag[.]commander[.]com", "tagcommander[.]com"],
    # AT Internet / Piano Analytics
    "Piano Analytics (AT Internet)":    ["atinternet[.]com", "piano[.]io/collect", "ati-host[.]net"],
    # Matomo / Piwik
    "Matomo":                           ["matomo[.]js", "piwik[.]js", "matomo[.]php", "piwik[.]php"],
    # Heap
    "Heap":                             ["heapanalytics[.]com", "cdn[.]heapanalytics[.]com"],
    # Mixpanel
    "Mixpanel":                         ["mixpanel[.]com/track", "cdn[.]mxpnl[.]com"],
    # mParticle
    "mParticle":                        ["mparticle[.]com/events"],
    # Amplitude
    "Amplitude":                        ["api[.]amplitude[.]com", "api2[.]amplitude[.]com"],
    # Snowplow
    "Snowplow":                         ["/com[.]snowplowanalytics", "snowplow[.]js"],
}

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────

# Paleta de colores Wormlytics
C_HEADER_DARK  = "1A1A2E"   # Azul muy oscuro — cabeceras principales
C_HEADER_MID   = "16213E"   # Azul oscuro — sub-cabeceras
C_ACCENT       = "0F3460"   # Azul medio — acento
C_HIGHLIGHT    = "E94560"   # Rojo/coral — highlights importantes
C_ROW_ALT      = "F0F4FF"   # Azul muy claro — filas alternas
C_WHITE        = "FFFFFF"
C_LIGHT_GRAY   = "F8F9FA"
C_BORDER       = "CCD6F6"

def hdr(text, bold=True, size=11, color=C_WHITE, bg=C_HEADER_DARK):
    """Crea una celda de cabecera estilizada."""
    return {
        "value": text,
        "font": Font(name="Arial", bold=bold, size=size, color=color),
        "fill": PatternFill("solid", fgColor=bg),
        "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "border": Border(
            left=Side(style="thin", color=C_BORDER),
            right=Side(style="thin", color=C_BORDER),
            top=Side(style="thin", color=C_BORDER),
            bottom=Side(style="thin", color=C_BORDER),
        )
    }

def cell_style(value, bold=False, bg=None, color="000000", wrap=False, align="left"):
    return {
        "value": value,
        "font": Font(name="Arial", bold=bold, size=10, color=color),
        "fill": PatternFill("solid", fgColor=bg) if bg else PatternFill(),
        "alignment": Alignment(horizontal=align, vertical="center", wrap_text=wrap),
        "border": Border(
            left=Side(style="hair", color="DDDDDD"),
            right=Side(style="hair", color="DDDDDD"),
            top=Side(style="hair", color="EEEEEE"),
            bottom=Side(style="hair", color="EEEEEE"),
        )
    }

def apply(ws, row, col, style_dict):
    c = ws.cell(row=row, column=col, value=style_dict["value"])
    for attr in ("font","fill","alignment","border"):
        if style_dict.get(attr):
            setattr(c, attr, style_dict[attr])
    return c

def _parse_ecommerce_events(pages_data: list) -> list:
    """
    Extrae y normaliza eventos de e-commerce de:
      - GA4 Enhanced Ecommerce (view_item, add_to_cart, begin_checkout, purchase, etc.)
      - UA Enhanced Ecommerce (detail, add, checkout, purchase)
      - Adobe Analytics products string (";product;category;qty;price;events")
      - Cualquier evento GTM con ecommerce.items[] o ecommerce.products[]
    Devuelve lista de dicts normalizados con campos comunes.
    """
    GA4_EVENTS = {
        "view_item":          "Ver producto",
        "view_item_list":     "Ver listado",
        "select_item":        "Seleccionar producto",
        "add_to_cart":        "Añadir al carrito",
        "remove_from_cart":   "Quitar del carrito",
        "view_cart":          "Ver carrito",
        "begin_checkout":     "Inicio checkout",
        "add_payment_info":   "Info de pago",
        "add_shipping_info":  "Info de envío",
        "purchase":           "Compra completada",
        "refund":             "Reembolso",
        "view_promotion":     "Ver promoción",
        "select_promotion":   "Seleccionar promoción",
        # UA legacy
        "productDetail":      "Ver producto (UA)",
        "addToCart":          "Añadir carrito (UA)",
        "removeFromCart":     "Quitar carrito (UA)",
        "checkout":           "Checkout (UA)",
        "purchase":           "Compra (UA)",
        # Genéricos que suelen tener datos de producto
        "product_impression": "Impresión producto",
        "product_click":      "Click producto",
        "checkout_progress":  "Progreso checkout",
        "order_complete":     "Pedido completado",
        "transaction":        "Transacción",
    }

    rows = []

    def extract_items(payload: dict) -> list:
        """Extrae el array de items/products de un payload de ecommerce."""
        ec = payload.get("ecommerce", {})
        if not isinstance(ec, dict):
            return []
        # GA4
        items = ec.get("items", [])
        if items:
            return items
        # UA: detail.products, add.products, checkout.products, purchase.products
        for key in ("detail","add","remove","checkout","purchase","impressions","click"):
            node = ec.get(key, {})
            if isinstance(node, dict):
                prods = node.get("products", node.get("items", []))
                if prods:
                    return prods
        return []

    def extract_transaction(payload: dict) -> dict:
        """Extrae datos de transacción (purchase) del payload."""
        ec = payload.get("ecommerce", {})
        if not isinstance(ec, dict):
            return {}
        # GA4
        txn = {
            "transaction_id": ec.get("transaction_id",""),
            "value":          ec.get("value",""),
            "tax":            ec.get("tax",""),
            "shipping":       ec.get("shipping",""),
            "currency":       ec.get("currency",""),
            "coupon":         ec.get("coupon",""),
        }
        # UA: purchase.actionField
        af = ec.get("purchase", {}).get("actionField", {})
        if af:
            txn["transaction_id"] = txn["transaction_id"] or af.get("id","")
            txn["value"]          = txn["value"]          or af.get("revenue","")
            txn["tax"]            = txn["tax"]             or af.get("tax","")
            txn["shipping"]       = txn["shipping"]        or af.get("shipping","")
            txn["coupon"]         = txn["coupon"]          or af.get("coupon","")
        # checkout actionField step
        cf = ec.get("checkout", {}).get("actionField", {})
        if cf:
            txn["checkout_step"]   = cf.get("step","")
            txn["checkout_option"] = cf.get("option","")
        return txn

    def safe(v, maxlen=120):
        if v is None: return ""
        if isinstance(v, (dict, list)): return json.dumps(v, ensure_ascii=False)[:maxlen]
        return str(v)[:maxlen]

    seen_keys = set()

    for pg in pages_data:
        page_title = pg.get("title","")[:70]
        page_url   = pg.get("url","")[:120]

        # ── GA4 / GTM Enhanced Ecommerce ─────────────────────────────────────
        dl = pg.get("google", {}).get("dataLayer", [])
        for item in dl:
            if not isinstance(item, dict):
                continue
            event_name = item.get("event","")
            if event_name not in GA4_EVENTS and "ecommerce" not in item:
                continue

            event_label = GA4_EVENTS.get(event_name, event_name)
            txn    = extract_transaction(item)
            items  = extract_items(item)
            ec_raw = item.get("ecommerce", {})

            if items:
                for prod in items:
                    if not isinstance(prod, dict):
                        continue
                    dedup_key = (event_name,
                                 prod.get("item_id","") or prod.get("id",""),
                                 page_url)
                    if dedup_key in seen_keys:
                        continue
                    seen_keys.add(dedup_key)

                    rows.append({
                        "evento":           event_name,
                        "evento_label":     event_label,
                        "fuente":           "GTM_dataLayer",
                        "item_id":          safe(prod.get("item_id") or prod.get("id","")),
                        "item_name":        safe(prod.get("item_name") or prod.get("name","")),
                        "item_brand":       safe(prod.get("item_brand") or prod.get("brand","")),
                        "item_category":    safe(prod.get("item_category") or prod.get("category","")),
                        "item_category2":   safe(prod.get("item_category2","")),
                        "item_variant":     safe(prod.get("item_variant") or prod.get("variant","")),
                        "item_list_name":   safe(prod.get("item_list_name") or prod.get("list","")),
                        "item_list_id":     safe(prod.get("item_list_id","")),
                        "index":            safe(prod.get("index") or prod.get("position","")),
                        "price":            safe(prod.get("price","")),
                        "quantity":         safe(prod.get("quantity","")),
                        "coupon":           safe(prod.get("coupon","") or txn.get("coupon","")),
                        "discount":         safe(prod.get("discount","")),
                        "affiliation":      safe(prod.get("affiliation","") or ec_raw.get("affiliation","")),
                        "transaction_id":   safe(txn.get("transaction_id","")),
                        "revenue_total":    safe(txn.get("value","")),
                        "tax":              safe(txn.get("tax","")),
                        "shipping":         safe(txn.get("shipping","")),
                        "currency":         safe(txn.get("currency","") or ec_raw.get("currency","")),
                        "checkout_step":    safe(txn.get("checkout_step","")),
                        "checkout_option":  safe(txn.get("checkout_option","")),
                        "promotion_id":     safe(ec_raw.get("promotion_id","") or ec_raw.get("creative","")),
                        "promotion_name":   safe(ec_raw.get("promotion_name","") or ec_raw.get("name","")),
                        "page_title":       page_title,
                        "page_url":         page_url,
                    })
            else:
                # Evento sin items pero con datos de ecommerce (p.ej. begin_checkout sin items)
                if event_name in GA4_EVENTS:
                    dedup_key = (event_name, "", page_url)
                    if dedup_key not in seen_keys:
                        seen_keys.add(dedup_key)
                        rows.append({
                            "evento":           event_name,
                            "evento_label":     event_label,
                            "fuente":           "GTM_dataLayer",
                            "item_id":          "",
                            "item_name":        "",
                            "item_brand":       "",
                            "item_category":    "",
                            "item_category2":   "",
                            "item_variant":     "",
                            "item_list_name":   "",
                            "item_list_id":     "",
                            "index":            "",
                            "price":            "",
                            "quantity":         "",
                            "coupon":           safe(txn.get("coupon","")),
                            "discount":         "",
                            "affiliation":      safe(ec_raw.get("affiliation","")),
                            "transaction_id":   safe(txn.get("transaction_id","")),
                            "revenue_total":    safe(txn.get("value","")),
                            "tax":              safe(txn.get("tax","")),
                            "shipping":         safe(txn.get("shipping","")),
                            "currency":         safe(txn.get("currency","") or ec_raw.get("currency","")),
                            "checkout_step":    safe(txn.get("checkout_step","")),
                            "checkout_option":  safe(txn.get("checkout_option","")),
                            "promotion_id":     "",
                            "promotion_name":   "",
                            "page_title":       page_title,
                            "page_url":         page_url,
                        })

        # ── Adobe Analytics products string ──────────────────────────────────
        # Formato: [Name];[SKU/Category];[Qty];[Price/TotalRevenue];[Events];[eVars/Props]
        s_obj    = pg.get("adobe", {}).get("s_object", {})
        products = s_obj.get("products","")
        aa_events = s_obj.get("events","")
        if products:
            for prod_str in products.split(","):
                parts = prod_str.strip().split(";")
                if len(parts) < 2:
                    continue
                dedup_key = ("aa_product", parts[1] if len(parts)>1 else "", page_url)
                if dedup_key in seen_keys:
                    continue
                seen_keys.add(dedup_key)
                # Inferir evento a partir de AA events string
                evt = "view_item"
                if aa_events:
                    if "purchase" in aa_events.lower() or "event12" in aa_events.lower():
                        evt = "purchase"
                    elif "scAdd" in aa_events or "event11" in aa_events:
                        evt = "add_to_cart"
                    elif "scCheckout" in aa_events:
                        evt = "begin_checkout"
                rows.append({
                    "evento":           evt,
                    "evento_label":     GA4_EVENTS.get(evt, evt),
                    "fuente":           "Adobe_products_string",
                    "item_id":          "",
                    "item_name":        safe(parts[0]) if len(parts)>0 else "",
                    "item_brand":       "",
                    "item_category":    safe(parts[1]) if len(parts)>1 else "",
                    "item_category2":   "",
                    "item_variant":     "",
                    "item_list_name":   "",
                    "item_list_id":     "",
                    "index":            "",
                    "price":            safe(parts[3]) if len(parts)>3 else "",
                    "quantity":         safe(parts[2]) if len(parts)>2 else "",
                    "coupon":           "",
                    "discount":         "",
                    "affiliation":      "",
                    "transaction_id":   safe(s_obj.get("purchaseID","") or s_obj.get("transactionID","")),
                    "revenue_total":    safe(parts[3]) if len(parts)>3 else "",
                    "tax":              "",
                    "shipping":         "",
                    "currency":         safe(s_obj.get("currencyCode","")),
                    "checkout_step":    "",
                    "checkout_option":  "",
                    "promotion_id":     "",
                    "promotion_name":   "",
                    "page_title":       page_title,
                    "page_url":         page_url,
                })

    return rows


def build_excel(scan_data: dict, pages_data: list, network_hits: dict, output_path: str):
    """
    Construye el Excel con 6 pestañas:
      1. RESUMEN        — metadata + herramientas detectadas
      2. VARIABLES_AA   — todas las variables de Adobe Analytics
      3. DATALAYER_GTM  — todos los eventos del dataLayer de GTM
      4. EVENTOS_RT     — eventos capturados en runtime (push/tl)
      5. PAGINAS        — variables por página visitada
      6. ECOMMERCE      — eventos de producto/carrito/checkout/compra (GA4 EE + UA EE + AA products)
    """
    wb = openpyxl.Workbook()

    # ── Hoja 1: RESUMEN ───────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "RESUMEN"
    ws1.sheet_view.showGridLines = False
    ws1.freeze_panes = "A2"

    # Título
    ws1.merge_cells("A1:F1")
    title_cell = ws1["A1"]
    title_cell.value = f"🪱 WORMLYTICS — Análisis de Analytics  |  {scan_data['domain']}  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    title_cell.font = Font(name="Arial", bold=True, size=14, color=C_WHITE)
    title_cell.fill = PatternFill("solid", fgColor=C_HEADER_DARK)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32

    # Metadata
    meta_rows = [
        ("URL analizada",    scan_data.get("url", "")),
        ("Dominio",          scan_data.get("domain", "")),
        ("Título de página", scan_data.get("title", "")),
        ("Fecha de escaneo", scan_data.get("scan_time", "")),
        ("Páginas visitadas",str(scan_data.get("pages_visited", 1))),
        ("Duración (s)",     str(scan_data.get("duration_s", ""))),
        ("Versión Wormlytics","5.5"),
    ]
    ws1.merge_cells("A2:B2")
    ws1["A2"].value = "METADATA"
    ws1["A2"].font = Font(name="Arial", bold=True, size=10, color=C_WHITE)
    ws1["A2"].fill = PatternFill("solid", fgColor=C_ACCENT)
    ws1["A2"].alignment = Alignment(horizontal="center")

    for i, (k, v) in enumerate(meta_rows, start=3):
        ws1.cell(row=i, column=1, value=k).font = Font(name="Arial", bold=True, size=10)
        ws1.cell(row=i, column=1).fill = PatternFill("solid", fgColor=C_ROW_ALT)
        ws1.cell(row=i, column=2, value=v).font = Font(name="Arial", size=10)

    # Herramientas detectadas
    start_tool = len(meta_rows) + 5
    ws1.merge_cells(f"A{start_tool}:F{start_tool}")
    ws1[f"A{start_tool}"].value = "HERRAMIENTAS DE ANALYTICS DETECTADAS"
    ws1[f"A{start_tool}"].font = Font(name="Arial", bold=True, size=11, color=C_WHITE)
    ws1[f"A{start_tool}"].fill = PatternFill("solid", fgColor=C_HIGHLIGHT)
    ws1[f"A{start_tool}"].alignment = Alignment(horizontal="center")

    tool_hdr_row = start_tool + 1
    for col, txt in enumerate(["Herramienta", "Tipo", "Beacons capturados", "Report Suite / Container"], 1):
        apply(ws1, tool_hdr_row, col, hdr(txt, bg=C_HEADER_MID))

    tools = scan_data.get("tools_detected", [])
    for i, tool in enumerate(tools, start=tool_hdr_row+1):
        tipo = ("Adobe Analytics" if "Adobe" in tool else
                "Google Analytics/GTM" if "Google" in tool else "Otro")
        beacons = sum(len(v) for k, v in network_hits.items() if k.split(" ")[0] in tool)
        rs = scan_data.get("report_suite", "") if "Adobe" in tool else \
             ", ".join(scan_data.get("gtm_containers", []))
        bg = C_ROW_ALT if i % 2 == 0 else C_WHITE
        for col, val in enumerate([tool, tipo, str(beacons), rs], 1):
            apply(ws1, i, col, cell_style(val, bg=bg))

    # Anchos
    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 55
    ws1.column_dimensions["C"].width = 20
    ws1.column_dimensions["D"].width = 20
    ws1.column_dimensions["E"].width = 20
    ws1.column_dimensions["F"].width = 30

    # ── Hoja 2: VARIABLES_AA ─────────────────────────────────────────────────
    # Estructura: UNA FILA POR VARIABLE, todos sus valores distintos en columnas
    # Variable | Tipo | Categoría SDR | Nº valores | Valor 1 | Página 1 | Valor 2 | Página 2 | ...
    ws2 = wb.create_sheet("VARIABLES_AA")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = "D3"   # freeze Variable+Tipo+Categoría, scroll horizontal en valores

    def classify_var(k):
        if k.startswith("eVar"):   return "eVar (Conversion Variable)"
        if k.startswith("prop"):   return "prop (Traffic Variable)"
        if k == "events":          return "Events"
        if k == "pageName":        return "Page Name (clave)"
        if k in ("channel","hier1","hier2","hier3"): return "Navegación / Jerarquía"
        if k in ("campaign",):     return "Marketing / Campaña"
        if k in ("products","purchaseID","transactionID","currencyCode"): return "E-commerce"
        if k in ("visitorID",):    return "Identificación de usuario"
        return "Variable estándar"

    # ── 1. Recopilar: variable → lista ordenada de {value, title, url} únicos ─
    from collections import defaultdict, OrderedDict
    # var_data[var][col_idx] = value
    # col_idx 0..n_pages-1 = páginas reales, n_pages+ = beacons de red
    n_pages = len(pages_data)
    var_data = defaultdict(dict)   # var -> {col_idx: value}

    def store_var(k, v, col_idx):
        if not v or k in ("all_s_snapshots", "s_snapshots_count"): return
        vstr = str(v)[:150]
        if col_idx not in var_data[k]:   # first value wins for each col slot
            var_data[k][col_idx] = vstr

    # ── Páginas reales ────────────────────────────────────────────────────────
    for pg_idx, pg in enumerate(pages_data):
        adobe = pg.get("adobe", {})

        # s_object
        for k, v in adobe.get("s_object", {}).items():
            store_var(k, v, pg_idx)

        # s.t() snapshots
        for snap in adobe.get("all_s_snapshots", []):
            if isinstance(snap, dict):
                for k, v in snap.items():
                    store_var(k, v, pg_idx)

        # RT intercepts
        for evt in pg.get("runtime_events", []):
            if isinstance(evt, dict) and evt.get("source","").startswith("AA_s."):
                snap = evt.get("data", {}).get("s_snapshot", {})
                if isinstance(snap, dict):
                    for k, v in snap.items():
                        store_var(k, v, pg_idx)

        # Tealium utag_data / data_layer
        teal = pg.get("tealium", {})
        for dl_key in ("utag_data", "data_layer"):
            dl_obj = teal.get(dl_key)
            if isinstance(dl_obj, dict):
                for k, v in dl_obj.items():
                    store_var(k, v, pg_idx)

        # Commanders Act
        tc = (pg.get("commanders_act") or {}).get("vars")
        if isinstance(tc, dict):
            for k, v in tc.items():
                store_var(k, v, pg_idx)

    # ── Beacons de red ────────────────────────────────────────────────────────
    # Usar hit["vars"] directamente — ya parseado en _on_request
    beacon_col = n_pages
    for tool_name, hits in network_hits.items():
        for hit in hits:
            vars_dict = hit.get("vars", {})
            if not vars_dict:
                # backward compat: check old keys
                vars_dict = hit.get("aa_vars") or hit.get("alloy_vars") or                             hit.get("tealium_vars") or {}
            # Remove internal metadata keys
            clean = {k: v for k, v in vars_dict.items()
                     if v and not k.startswith("_") and k not in ("clientId","sessionId")}
            if clean:
                for k, v in clean.items():
                    store_var(k, v, beacon_col)
                beacon_col += 1

    n_beacon_cols  = beacon_col - n_pages
    total_cols_data = n_pages + n_beacon_cols









    # ── 2. Ordenar variables ────────────────────────────────────────────────
    def var_sort_key(k):
        if k.startswith("eVar"):
            try:    return (0, int(k[4:]), k)
            except: return (0, 9999, k)
        if k.startswith("prop"):
            try:    return (1, int(k[4:]), k)
            except: return (1, 9999, k)
        return (2, 0, k)

    sorted_vars = sorted(var_data.keys(), key=var_sort_key)

    # ── 3. Cabeceras ─────────────────────────────────────────────────────────
    # Compactamos los valores: una columna por valor único (en orden de aparición)
    # Esto evita tener docenas de "Pág N" vacías cuando los datos vienen de beacons.

    # Construir lista compacta de valores por variable
    # var_data[var][col_idx] = value (original)
    # var_compact[var] = [value1, value2, ...] (compactado, sin huecos)
    var_compact = {}
    for var in var_data:
        # Sort col_idx so páginas (0..n_pages-1) vienen antes que beacons (n_pages..)
        sorted_slots = sorted(var_data[var].keys())
        var_compact[var] = [var_data[var][s] for s in sorted_slots]

    # max número de valores a mostrar por variable (caps razonable)
    max_vals_per_var = max((len(v) for v in var_compact.values()), default=1)
    max_cols  = min(max_vals_per_var, 50)
    FIXED     = ["Variable", "Tipo", "Categoría SDR", "Nº valores"]
    total_c   = len(FIXED) + max_cols

    ws2.merge_cells(f"A1:{get_column_letter(total_c)}1")
    ws2["A1"].value = (
        f"Analytics — Variables capturadas  |  "
        f"{len(sorted_vars)} variables únicas  ·  {n_pages} páginas analizadas  ·  "
        f"{n_beacon_cols} beacons de red  ·  (una columna por valor capturado, sin huecos)"
    )
    ws2["A1"].font      = Font(name="Arial", bold=True, size=13, color=C_WHITE)
    ws2["A1"].fill      = PatternFill("solid", fgColor="CC0000")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    for col, txt in enumerate(FIXED, 1):
        apply(ws2, 2, col, hdr(txt, bg=C_HEADER_DARK))

    # Headers: Valor 1, Valor 2, ... (sin distinguir página/beacon — compactado)
    for n in range(1, max_cols + 1):
        col   = len(FIXED) + n
        color = "1B5E20" if n % 2 == 1 else "2E7D32"
        apply(ws2, 2, col, hdr(f"Valor {n}", bg=color, size=9))
    ws2.row_dimensions[2].height = 24

    # ── 4. Filas de datos ───────────────────────────────────────────────────
    for row_i, var in enumerate(sorted_vars, start=3):
        values  = var_compact.get(var, [])
        n_hits  = len(values)
        tipo    = "eVar" if var.startswith("eVar") else "prop" if var.startswith("prop") else "Estándar"
        categ   = classify_var(var)
        bg      = C_ROW_ALT if row_i % 2 == 0 else C_WHITE

        apply(ws2, row_i, 1, cell_style(var,         bold=True, bg=bg))
        apply(ws2, row_i, 2, cell_style(tipo,        bg=bg))
        apply(ws2, row_i, 3, cell_style(categ,       bg=bg))
        apply(ws2, row_i, 4, cell_style(str(n_hits), bg=bg, align="center",
                                         bold=True,
                                         color="1B5E20" if n_hits > 2 else "000000"))

        for slot, val in enumerate(values[:max_cols]):
            if val:
                c      = len(FIXED) + slot + 1
                val_bg = "F1F8E9" if slot % 2 == 0 else "E8F5E9"
                apply(ws2, row_i, c, cell_style(str(val)[:200], bg=val_bg, wrap=True))

    # ── 5. Anchos ────────────────────────────────────────────────────────────
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 8
    ws2.column_dimensions["C"].width = 26
    ws2.column_dimensions["D"].width = 10
    for n in range(1, max_cols + 1):
        ws2.column_dimensions[get_column_letter(len(FIXED) + n)].width = 28

    # ── Hoja 3: DATALAYER_GTM ────────────────────────────────────────────────
    ws3 = wb.create_sheet("DATALAYER_TMS")
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "A3"

    ws3.merge_cells("A1:G1")
    ws3["A1"].value = "TMS Data Layer — Eventos de GTM dataLayer, Tealium utag.data y otros TMS"
    ws3["A1"].font = Font(name="Arial", bold=True, size=13, color=C_WHITE)
    ws3["A1"].fill = PatternFill("solid", fgColor="0F9D58")   # verde Google
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    gtm_cols = ["Nombre del evento", "TMS", "Fuente (Página)", "Payload completo (JSON)", "Clave 1", "Valor 1", "Clave 2", "Valor 2"]
    for col, txt in enumerate(gtm_cols, 1):
        apply(ws3, 2, col, hdr(txt, bg=C_HEADER_DARK))

    # Recopilar eventos de TODOS los TMS: GTM dataLayer + Tealium utag.data + RT events
    all_gtm = []
    for pg in pages_data:
        page_title = pg.get("title","")[:60]

        # GTM dataLayer
        for item in pg.get("google", {}).get("dataLayer", []):
            if isinstance(item, dict):
                all_gtm.append({"event": item.get("event","(no event)"),
                                 "tms": "GTM", "page": page_title, "payload": item})

        # Tealium utag.data
        utag_data_obj = pg.get("tealium", {}).get("utag_data") or pg.get("tealium", {}).get("data_layer")
        if isinstance(utag_data_obj, dict):
            all_gtm.append({"event": utag_data_obj.get("tealium_event", "utag_data_snapshot"),
                             "tms": "Tealium", "page": page_title, "payload": utag_data_obj})

        # Extra data layers (digitalData, utag_data top-level, etc.)
        for dl_name, dl_val in pg.get("extra_datalayers", {}).items():
            if isinstance(dl_val, dict):
                all_gtm.append({"event": f"{dl_name}_snapshot", "tms": dl_name,
                                 "page": page_title, "payload": dl_val})
            elif isinstance(dl_val, list):
                for item in dl_val:
                    if isinstance(item, dict):
                        all_gtm.append({"event": item.get("event", f"{dl_name}_event"),
                                         "tms": dl_name, "page": page_title, "payload": item})

        # RT events from ALL TMS (GTM + Tealium + Segment + Piano)
        for evt in pg.get("runtime_events", []):
            if isinstance(evt, dict):
                src_evt = evt.get("source","")
                if any(x in src_evt for x in ["Tealium_","Segment_","Piano_","AA_"]):
                    data = evt.get("data", {})
                    all_gtm.append({"event": data.get("call","") or data.get("event","") or src_evt,
                                     "tms": src_evt.split("_")[0], "page": page_title,
                                     "payload": data})

    # Deduplicar por evento + payload
    seen = set()
    unique_gtm = []
    for g in all_gtm:
        key = g["event"] + json.dumps(g["payload"], sort_keys=True, ensure_ascii=False)[:80]
        if key not in seen:
            seen.add(key)
            unique_gtm.append(g)

    for i, g in enumerate(unique_gtm, start=3):
        bg = C_ROW_ALT if i % 2 == 0 else C_WHITE
        # Color-code by TMS
        tms = g.get("tms","GTM")
        tms_color = {"GTM":"E8F5E9","Tealium":"E3F2FD","Segment":"FFF3E0",
                     "Piano":"F3E5F5","Matomo":"FBE9E7"}.get(tms, C_WHITE)
        if i % 2 == 0: tms_color = C_ROW_ALT
        payload = g["payload"]
        keys    = [k for k in payload.keys() if k not in ("event","gtm.uniqueEventId","gtm.start")]
        k1, v1  = (keys[0], str(payload[keys[0]])[:60]) if len(keys) > 0 else ("","")
        k2, v2  = (keys[1], str(payload[keys[1]])[:60]) if len(keys) > 1 else ("","")
        payload_str = json.dumps(payload, ensure_ascii=False)[:300]
        for col, val in enumerate([g["event"], tms, g["page"], payload_str, k1, v1, k2, v2], 1):
            apply(ws3, i, col, cell_style(val, bold=(col==1), bg=tms_color, wrap=(col==4)))

    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 12
    ws3.column_dimensions["C"].width = 30
    ws3.column_dimensions["D"].width = 55
    ws3.column_dimensions["E"].width = 20
    ws3.column_dimensions["F"].width = 25
    ws3.column_dimensions["G"].width = 20
    ws3.column_dimensions["H"].width = 25

    # ── Hoja 4: EVENTOS_RT ────────────────────────────────────────────────────
    ws4 = wb.create_sheet("EVENTOS_RT")
    ws4.sheet_view.showGridLines = False
    ws4.freeze_panes = "A3"

    ws4.merge_cells("A1:F1")
    ws4["A1"].value = "Eventos capturados en Runtime (dataLayer.push interceptado + s.t / s.tl + beacons fetch/XHR)"
    ws4["A1"].font = Font(name="Arial", bold=True, size=13, color=C_WHITE)
    ws4["A1"].fill = PatternFill("solid", fgColor="7B2FBE")   # púrpura
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 28

    rt_cols = ["Timestamp", "Fuente / Tipo", "Nombre del evento", "Payload / Datos", "Página", "URL de la página"]
    for col, txt in enumerate(rt_cols, 1):
        apply(ws4, 2, col, hdr(txt, bg=C_HEADER_DARK))

    all_rt = []
    for pg in pages_data:
        for evt in pg.get("runtime_events", []):
            all_rt.append({**evt, "page_title": pg.get("title",""), "page_url": pg.get("url","")})
        for bcn in pg.get("runtime_beacons", []):
            all_rt.append({"ts": bcn.get("ts",""), "source": f"beacon/{bcn.get('type','')}", "data": {"url": bcn.get("url","")},
                           "page_title": pg.get("title",""), "page_url": pg.get("url","")})

    for i, evt in enumerate(all_rt, start=3):
        bg = C_ROW_ALT if i % 2 == 0 else C_WHITE
        ts_short = str(evt.get("ts",""))[-12:-4] if evt.get("ts") else ""
        src      = evt.get("source", "")
        name     = evt.get("data", {}).get("event", "") if isinstance(evt.get("data"),dict) else ""
        payload  = json.dumps(evt.get("data", {}), ensure_ascii=False)[:250]
        for col, val in enumerate([ts_short, src, name, payload,
                                    evt.get("page_title","")[:60], evt.get("page_url","")[:100]], 1):
            apply(ws4, i, col, cell_style(val, bg=bg, wrap=(col==4)))

    ws4.column_dimensions["A"].width = 12
    ws4.column_dimensions["B"].width = 22
    ws4.column_dimensions["C"].width = 28
    ws4.column_dimensions["D"].width = 55
    ws4.column_dimensions["E"].width = 35
    ws4.column_dimensions["F"].width = 45

    # ── Hoja 5: PAGINAS ───────────────────────────────────────────────────────
    ws5 = wb.create_sheet("PAGINAS")
    ws5.sheet_view.showGridLines = False
    ws5.freeze_panes = "A3"

    ws5.merge_cells("A1:H1")
    ws5["A1"].value = "Páginas visitadas — Variables por página"
    ws5["A1"].font = Font(name="Arial", bold=True, size=13, color=C_WHITE)
    ws5["A1"].fill = PatternFill("solid", fgColor=C_ACCENT)
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 28

    pg_cols = ["#", "Título de página", "URL", "pageName (AA)", "channel (AA)",
               "pageType (AA)", "Eventos GTM", "Variables AA populadas"]
    for col, txt in enumerate(pg_cols, 1):
        apply(ws5, 2, col, hdr(txt, bg=C_HEADER_DARK))

    for i, pg in enumerate(pages_data, start=3):
        s_obj = pg.get("adobe", {}).get("s_object", {})
        dl    = pg.get("google", {}).get("dataLayer", [])
        gtm_events = [x.get("event","") for x in dl if isinstance(x,dict) and "event" in x]
        bg = C_ROW_ALT if i % 2 == 0 else C_WHITE
        vals = [
            str(i-2),
            pg.get("title","")[:80],
            pg.get("url","")[:120],
            s_obj.get("pageName",""),
            s_obj.get("channel",""),
            s_obj.get("pageType",""),
            ", ".join(set(gtm_events))[:80],
            str(len(s_obj))
        ]
        for col, val in enumerate(vals, 1):
            bold = col == 1
            apply(ws5, i, col, cell_style(val, bold=bold, bg=bg))

    ws5.column_dimensions["A"].width = 5
    ws5.column_dimensions["B"].width = 45
    ws5.column_dimensions["C"].width = 55
    ws5.column_dimensions["D"].width = 35
    ws5.column_dimensions["E"].width = 20
    ws5.column_dimensions["F"].width = 18
    ws5.column_dimensions["G"].width = 40
    ws5.column_dimensions["H"].width = 22

    # ── Hoja 6: ECOMMERCE ────────────────────────────────────────────────────
    ws6 = wb.create_sheet("ECOMMERCE")
    ws6.sheet_view.showGridLines = False
    ws6.freeze_panes = "A3"

    ws6.merge_cells("A1:AC1")
    ws6["A1"].value = (
        "E-commerce — Eventos de producto, carrito, checkout y compra  "
        "(GA4 Enhanced Ecommerce + UA Enhanced Ecommerce + Adobe Analytics products string)"
    )
    ws6["A1"].font      = Font(name="Arial", bold=True, size=13, color=C_WHITE)
    ws6["A1"].fill      = PatternFill("solid", fgColor="E65100")   # naranja e-commerce
    ws6["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws6.row_dimensions[1].height = 28

    EC_COLS = [
        # Evento
        ("Evento (raw)",         "evento",          18),
        ("Evento (label)",       "evento_label",    22),
        ("Fuente",               "fuente",          20),
        # Producto
        ("Item ID / SKU",        "item_id",         18),
        ("Nombre del producto",  "item_name",       35),
        ("Marca",                "item_brand",      18),
        ("Categoría 1",          "item_category",   22),
        ("Categoría 2",          "item_category2",  22),
        ("Variante",             "item_variant",    16),
        ("Lista / Impresión",    "item_list_name",  22),
        ("ID Lista",             "item_list_id",    16),
        ("Posición",             "index",           10),
        ("Precio",               "price",           12),
        ("Cantidad",             "quantity",        10),
        ("Descuento",            "discount",        12),
        ("Cupón",                "coupon",          15),
        ("Afiliación",           "affiliation",     16),
        # Transacción
        ("Transaction ID",       "transaction_id",  22),
        ("Revenue total",        "revenue_total",   15),
        ("Tax",                  "tax",             10),
        ("Shipping",             "shipping",        12),
        ("Moneda",               "currency",        10),
        # Checkout
        ("Checkout step",        "checkout_step",   14),
        ("Checkout option",      "checkout_option", 16),
        # Promoción
        ("Promotion ID",         "promotion_id",    16),
        ("Promotion name",       "promotion_name",  22),
        # Contexto
        ("Página",               "page_title",      38),
        ("URL",                  "page_url",        45),
    ]

    # Grupo de colores por sección
    GROUP_COLORS = {
        "evento":          "1A237E",   # azul oscuro — evento
        "evento_label":    "1A237E",
        "fuente":          "1A237E",
        "item_id":         "1B5E20",   # verde oscuro — producto
        "item_name":       "1B5E20",
        "item_brand":      "1B5E20",
        "item_category":   "1B5E20",
        "item_category2":  "1B5E20",
        "item_variant":    "1B5E20",
        "item_list_name":  "1B5E20",
        "item_list_id":    "1B5E20",
        "index":           "1B5E20",
        "price":           "1B5E20",
        "quantity":        "1B5E20",
        "discount":        "1B5E20",
        "coupon":          "1B5E20",
        "affiliation":     "1B5E20",
        "transaction_id":  "B71C1C",   # rojo — transacción
        "revenue_total":   "B71C1C",
        "tax":             "B71C1C",
        "shipping":        "B71C1C",
        "currency":        "B71C1C",
        "checkout_step":   "4A148C",   # púrpura — checkout
        "checkout_option": "4A148C",
        "promotion_id":    "E65100",   # naranja — promo
        "promotion_name":  "E65100",
        "page_title":      "37474F",   # gris — contexto
        "page_url":        "37474F",
    }

    # Cabeceras con color por grupo
    for col, (label, field, _) in enumerate(EC_COLS, 1):
        bg = GROUP_COLORS.get(field, C_HEADER_DARK)
        apply(ws6, 2, col, hdr(label, bg=bg, size=10))
    ws6.row_dimensions[2].height = 36

    # Filas de datos
    ec_rows = _parse_ecommerce_events(pages_data)

    if ec_rows:
        for i, row in enumerate(ec_rows, start=3):
            bg = C_ROW_ALT if i % 2 == 0 else C_WHITE

            # Colorear filas de purchase en amarillo suave
            if row.get("evento") in ("purchase","transaction","order_complete"):
                bg = "FFFDE7"
            elif row.get("evento") in ("add_to_cart","addToCart"):
                bg = "E8F5E9"
            elif row.get("evento","").startswith("begin_checkout") or row.get("evento") == "checkout":
                bg = "EDE7F6"

            for col, (_, field, _) in enumerate(EC_COLS, 1):
                val  = row.get(field, "")
                bold = (field in ("evento_label", "item_name", "transaction_id", "revenue_total"))
                color = "000000"
                if field == "revenue_total" and val:
                    try:
                        float(str(val).replace(",",".").replace("€","").strip())
                        color = "1B5E20"   # verde si hay valor numérico
                    except ValueError:
                        pass
                apply(ws6, i, col, cell_style(val, bold=bold, bg=bg, color=color))
    else:
        ws6.merge_cells("A3:AC3")
        ws6["A3"].value = (
            "⚠  No se detectaron eventos de e-commerce en las páginas analizadas. "
            "Esto es normal si la web no tiene e-commerce o si los eventos se disparan "
            "solo en interacciones de usuario (añadir al carrito, checkout, compra)."
        )
        ws6["A3"].font      = Font(name="Arial", italic=True, size=10, color="666666")
        ws6["A3"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws6.row_dimensions[3].height = 48

    # Anchos de columna
    for col, (_, _, width) in enumerate(EC_COLS, 1):
        ws6.column_dimensions[get_column_letter(col)].width = width

    # ── Leyenda de colores en ECOMMERCE ──────────────────────────────────────
    legend_row = max(len(ec_rows) + 4, 5)
    ws6.cell(row=legend_row, column=1, value="Leyenda de colores de cabecera:").font = \
        Font(name="Arial", bold=True, size=9, color="666666")
    legends = [
        ("Azul oscuro", "1A237E", "Datos del evento"),
        ("Verde oscuro", "1B5E20", "Datos del producto"),
        ("Rojo",         "B71C1C", "Datos de transacción / purchase"),
        ("Púrpura",      "4A148C", "Datos de checkout"),
        ("Naranja",      "E65100", "Datos de promoción"),
        ("Gris",         "37474F", "Contexto de página"),
    ]
    for j, (name, color, desc) in enumerate(legends, 1):
        c = ws6.cell(row=legend_row + j, column=1, value=f"■ {name}: {desc}")
        c.font = Font(name="Arial", size=9, color=color, bold=True)

    wb.save(output_path)
    return output_path

# ─────────────────────────────────────────────────────────────────────────────
# WORMLYTICS BOT
# ─────────────────────────────────────────────────────────────────────────────



def _safe_json_default(obj):
    """Serializer seguro para tipos no serializables (datetime, bytes, sets, etc.)"""
    import datetime as dt
    if isinstance(obj, (dt.datetime, dt.date, dt.time)):
        return obj.isoformat()
    if isinstance(obj, bytes):
        return obj.decode("utf-8", errors="replace")
    if isinstance(obj, set):
        return list(obj)
    try:
        return str(obj)
    except Exception:
        return None


def _sanitize_for_json(data):
    """Recorre recursivamente y convierte tipos no serializables."""
    if isinstance(data, dict):
        return {k: _sanitize_for_json(v) for k, v in data.items()}
    if isinstance(data, (list, tuple)):
        return [_sanitize_for_json(i) for i in data]
    try:
        json.dumps(data)
        return data
    except (TypeError, ValueError):
        return str(data)




# ─────────────────────────────────────────────────────────────────────────────
# BEACON PARSERS  —  Lee directamente de los hits de red (fuente más fiable)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_aa_beacon(url: str) -> dict:
    """
    Adobe Analytics AppMeasurement — GET /b/ss/rsid/1/...?v1=eVar1&c1=prop1&pageName=...
    Mapeo completo de query params a nombres legibles.
    """
    from urllib.parse import urlparse, parse_qs, unquote_plus
    try:
        qs      = parse_qs(urlparse(url).query, keep_blank_values=False)
        out     = {}

        # eVars: v1..v250
        for i in range(1, 251):
            v = qs.get(f"v{i}", [None])[0]
            if v and v.strip():
                out[f"eVar{i}"] = unquote_plus(v)

        # props: c1..c75
        for i in range(1, 76):
            v = qs.get(f"c{i}", [None])[0]
            if v and v.strip():
                out[f"prop{i}"] = unquote_plus(v)

        # Variables estándar
        STANDARD = {
            "pageName": "pageName",  "ch":  "channel",    "server": "server",
            "h1": "hier1",           "h2":  "hier2",       "h3": "hier3",
            "h4": "hier4",           "h5":  "hier5",
            "v0": "campaign",        "pid": "pageURL",     "pidt": "pageType",
            "pe": "linkType",        "pev1":"linkURL",      "pev2":"linkName",
            "events":  "events",     "products": "products",
            "purchaseID":"purchaseID","transactionID":"transactionID",
            "zip":"zip",             "state":"state",       "cc":"currencyCode",
            "charSet":"charSet",     "s_account":"reportSuite",
        }
        for param, name in STANDARD.items():
            v = qs.get(param, [None])[0]
            if v and v.strip():
                out[name] = unquote_plus(v)

        # Report suite desde path: /b/ss/{rsid}/
        m = re.search(r"/b/ss/([^/]+)/", url)
        if m:
            out["_reportSuite"] = m.group(1)

        return out
    except Exception:
        return {}


def _parse_ga4_beacon(url: str, body: str = "") -> dict:
    """
    Google Analytics 4 — POST /g/collect?...
    Parámetros en la URL + body (ambos URL-encoded).
    """
    from urllib.parse import urlparse, parse_qs, unquote_plus
    try:
        out = {}

        def parse_params(raw: str):
            return parse_qs(raw, keep_blank_values=False)

        # Params from URL query string
        qs = parse_params(urlparse(url).query)
        # Params from body (if POST)
        if body:
            qs.update(parse_params(body))

        # Measurement ID / stream
        for k in ("tid", "measurement_id"):
            v = qs.get(k, [None])[0]
            if v: out["_measurementId"] = v

        # Client/session IDs
        for k, n in (("cid","clientId"),("sid","sessionId"),("uid","userId")):
            v = qs.get(k, [None])[0]
            if v: out[n] = v

        # Event name
        en = qs.get("en", [None])[0]
        if en: out["eventName"] = unquote_plus(en)

        # Page info
        dl = qs.get("dl", [None])[0]   # document location
        dt = qs.get("dt", [None])[0]   # document title
        dp = qs.get("dp", [None])[0]   # document path
        if dl: out["pageURL"]   = unquote_plus(dl)
        if dt: out["pageTitle"] = unquote_plus(dt)
        if dp: out["pagePath"]  = unquote_plus(dp)

        # Custom event params: ep.* (event params) and up.* (user params)
        for key, val_list in qs.items():
            val = val_list[0] if val_list else ""
            if not val: continue
            if key.startswith("ep."):
                out[key[3:]] = unquote_plus(val)   # e.g. ep.page_category → page_category
            elif key.startswith("up."):
                out[f"user.{key[3:]}"] = unquote_plus(val)
            elif key.startswith("pr") and "nm" in key:
                # Product name in ecommerce
                out[key] = unquote_plus(val)

        # Ecommerce: tr=revenue, tt=tax, ts=shipping
        for k, n in (("tr","revenue"),("tt","tax"),("ts","shipping"),("ti","transactionId")):
            v = qs.get(k, [None])[0]
            if v: out[n] = v

        return out
    except Exception:
        return {}


def _parse_alloy_beacon(url: str, body: str) -> dict:
    """
    Adobe Web SDK (Alloy) — POST /ee/v2/collect  body=JSON
    Formatos: [{events:[{xdm,data}]}]  |  [{xdm,data}]  |  {events:[...]}
    """
    try:
        payload = json.loads(body) if isinstance(body, str) and body.strip() else None
        if not payload: return {}
        out = {}

        def process_event(evt):
            if not isinstance(evt, dict): return
            aa = (evt.get("data") or {}).get("__adobe", {}).get("analytics", {})
            for k, v in aa.items():
                if v is not None and str(v).strip():
                    out[k] = str(v)[:200]
            _xdm_extract(evt.get("xdm") or {}, out)

        items = payload if isinstance(payload, list) else [payload]
        for item in items:
            if not isinstance(item, dict): continue
            if "events" in item and isinstance(item["events"], list):
                for evt in item["events"]:
                    process_event(evt)
            else:
                process_event(item)
        return out
    except Exception:
        return {}


def _xdm_extract(xdm: dict, out: dict):
    """Extrae solo los campos XDM relevantes para analytics (no metadata técnica)."""
    if not isinstance(xdm, dict): return

    # Page
    wp = (xdm.get("web") or {}).get("webPageDetails") or {}
    if wp.get("name"):  out.setdefault("pageName",  str(wp["name"])[:200])
    if wp.get("URL"):   out.setdefault("pageURL",   str(wp["URL"])[:200])
    if wp.get("server"):out.setdefault("server",    str(wp["server"])[:200])

    # Event type
    if xdm.get("eventType"):
        out.setdefault("xdm.eventType", str(xdm["eventType"]))

    # Commerce
    commerce = xdm.get("commerce") or {}
    order = commerce.get("order") or {}
    if order.get("purchaseID"):  out.setdefault("purchaseID",   str(order["purchaseID"]))
    if order.get("priceTotal"):  out.setdefault("revenue",      str(order["priceTotal"]))
    if order.get("currencyCode"):out.setdefault("currencyCode", str(order["currencyCode"]))

    # Products
    items = xdm.get("productListItems") or []
    if items:
        parts = []
        for item in items[:10]:
            if isinstance(item, dict):
                parts.append(f"{item.get('name','')};{item.get('SKU','')};{item.get('quantity','')};{item.get('priceTotal','')}")
        if parts: out.setdefault("products", " | ".join(parts)[:300])

    # Marketing
    mkt = xdm.get("marketing") or {}
    if mkt.get("trackingCode"): out.setdefault("campaign", str(mkt["trackingCode"]))

    # Search
    srch = xdm.get("search") or {}
    if srch.get("keywords"): out.setdefault("searchKeywords", str(srch["keywords"]))



class Wormlytics:
    def __init__(self, url, output_file=None, cdp_port=CDP_PORT,
                 nav_duration=NAV_DURATION, manual_mode=False, scope_path=None):
        self.url          = url
        self.output_file  = output_file
        self.cdp_port     = cdp_port
        self.nav_duration = nav_duration
        self.manual_mode  = manual_mode
        # scope_path: si se define, solo se navegan URLs que contengan este path
        from urllib.parse import urlparse as _up
        self.scope_path   = scope_path or _up(url).path.rstrip('/') or None
        self.domain       = urlparse(url).netloc.replace("www.","")
        self.network_hits = {}
        self.pages_data   = []          # datos extraídos por página
        self.visited_urls = set()
        self.step_count   = 0
        self.start_time   = None

    def step(self, icon, msg, detail=None, style="bold cyan"):
        self.step_count += 1
        elapsed = f"{time.time()-self.start_time:.1f}s" if self.start_time else "—"
        ts = datetime.now().strftime("%H:%M:%S")
        line = Text()
        line.append(f"  [{ts}] ", "dim")
        line.append(f"[{self.step_count:02d}] ", "bold white")
        line.append(f"{icon} ", "bold yellow")
        line.append(msg, style)
        if detail: line.append(f"  → {detail}", "dim")
        line.append(f"  ({elapsed})", "dim green")
        console.print(line)

    def warn(self, m):    console.print(f"  [yellow]  ⚠  {m}[/yellow]")
    def success(self, m): console.print(f"  [green]  ✓  {m}[/green]")
    def error(self, m):   console.print(f"  [red]  ✗  {m}[/red]")
    def info(self, m):    console.print(f"  [dim]  ℹ  {m}[/dim]")

    def _on_request(self, req):
        for tool, patterns in ANALYTICS_PATTERNS.items():
            for pat in patterns:
                if re.search(pat, req.url, re.IGNORECASE):
                    body = ""
                    try:
                        body = req.post_data or ""
                    except Exception:
                        pass

                    hit = {
                        "url":    req.url,
                        "method": req.method,
                        "ts":     datetime.now().isoformat(),
                        "vars":   {}   # parsed analytics variables
                    }

                    # Adobe Analytics AppMeasurement — GET /b/ss/
                    if "b/ss/" in req.url:
                        hit["vars"] = _parse_aa_beacon(req.url)
                        hit["source"] = "AA_beacon"

                    # Adobe Web SDK (Alloy) — POST /ee/v2/collect
                    elif tool == "Adobe Web SDK (alloy)" and body:
                        hit["vars"] = _parse_alloy_beacon(req.url, body)
                        hit["source"] = "alloy_beacon"

                    # Google Analytics 4 — POST /g/collect
                    elif "google-analytics.com/g/collect" in req.url or                          "analytics.google.com/g/collect" in req.url:
                        hit["vars"] = _parse_ga4_beacon(req.url, body)
                        hit["source"] = "ga4_beacon"

                    # Tealium collect
                    elif tool == "Tealium AudienceStream" and body:
                        try:
                            hit["vars"] = json.loads(body)
                            hit["source"] = "tealium_beacon"
                        except Exception:
                            pass

                    else:
                        hit["source"] = tool

                    self.network_hits.setdefault(tool, []).append(hit)
                    return

    async def inject_and_extract(self, page) -> dict:
        """Extrae datos de la página. Los interceptores ya están activos via add_init_script."""
        # Re-patch s object por si Adobe cargó después del init script
        try:
            await page.evaluate("""
            () => {
                if (typeof s !== 'undefined' && !s.__wl_patched) {
                    const snapS = () => {
                        const so = window.s, snap = {};
                        try {
                            for(let i=1;i<=250;i++){const k=`eVar${i}`;if(so[k]!==undefined&&so[k]!==''&&so[k]!==null)snap[k]=so[k];}
                            for(let i=1;i<=75;i++){const k=`prop${i}`;if(so[k]!==undefined&&so[k]!==''&&so[k]!==null)snap[k]=so[k];}
                            ['pageName','pageURL','channel','server','hier1','hier2','hier3','campaign',
                             'products','purchaseID','visitorID','charSet','pageType','transactionID',
                             'country','linkName','linkType','currencyCode','events'].forEach(k=>{if(so[k])snap[k]=so[k];});
                        } catch(e){}
                        return snap;
                    };
                    ['t','tl'].forEach(m => {
                        if(typeof s[m]==='function'){
                            const orig=s[m].bind(s);
                            s[m]=function(...args){
                                const snap=snapS();
                                if(!window.__wl_s_snapshots)window.__wl_s_snapshots=[];
                                window.__wl_s_snapshots.push(snap);
                                if(!window.__wl_events)window.__wl_events=[];
                                window.__wl_events.push({source:`AA_s.${m}`,ts:new Date().toISOString(),data:{method:m,s_snapshot:snap}});
                                return orig(...args);
                            };
                        }
                    });
                    s.__wl_patched=true;
                }
                // Also patch dataLayer if not done yet
                if(typeof dataLayer!=='undefined'&&!dataLayer.__wl_patched){
                    const orig=dataLayer.push.bind(dataLayer);
                    dataLayer.push=function(...args){
                        args.forEach(evt=>{
                            try{if(!window.__wl_events)window.__wl_events=[];
                            window.__wl_events.push({source:'GTM',ts:new Date().toISOString(),data:JSON.parse(JSON.stringify(evt))});}catch(e){}
                        });
                        return orig(...args);
                    };
                    dataLayer.__wl_patched=true;
                }
            }
            """)
        except Exception:
            pass
        # Small wait for any pending analytics calls to fire
        await page.wait_for_timeout(random.randint(400, 700))
        try:
            data = await page.evaluate(JS_EXTRACT)
        except Exception:
            data = {"url": page.url, "title": "", "adobe": {}, "google": {},
                    "gtm_containers": [], "runtime_events": [], "runtime_beacons": []}
        return data

    async def scroll_page(self, page):
        """Scroll rápido — activa lazy-load y eventos sin desperdiciar tiempo."""
        try:
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight * 0.5)")
            await asyncio.sleep(random.uniform(0.12, 0.22))
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            await asyncio.sleep(random.uniform(0.12, 0.20))
            await page.evaluate("window.scrollTo(0, 0)")
        except Exception:
            pass

    async def get_internal_links(self, page) -> list:
        """Obtiene links internos — filtrados al scope_path si está definido."""
        try:
            links = await page.evaluate(JS_GET_LINKS, self.domain)
            if self.scope_path:
                links = [l for l in links if self.scope_path in l]
            return links
        except Exception:
            return []

    _URL_BLACKLIST = re.compile(
        r'/secure/|/login|/signin|/sign-in|/logout|/account/|/my-account'
        r'|/checkout|/cart/|/basket/|/order/|/payment|/confirmation'
        r'|/register|/signup|/sign-up|/auth/'
        r'|[?]token=|[?]session=',
        re.IGNORECASE
    )

    async def navigate_page(self, page, url: str) -> bool:
        """Navega a una URL, extrae datos y continúa."""
        if url in self.visited_urls:
            return False
        # Saltar páginas de login / área privada / transacciones
        if self._URL_BLACKLIST.search(url):
            self.visited_urls.add(url)   # marcar como visitada para no reencolar
            return False
        self.visited_urls.add(url)
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=15000)
            await page.wait_for_timeout(random.randint(1200, 1800))
            title = await page.title()
            self.info(f"→ {title[:60]}  [{url[:70]}]")
            await self.scroll_page(page)
            data = await self.inject_and_extract(page)
            if data:
                self.pages_data.append(data)
            return True
        except Exception as e:
            err = str(e)
            if "crashed" in err.lower() or "Page.goto" in err:
                self.warn(f"⚠ Page crashed en {url[:50]} — esperando recuperación...")
                await asyncio.sleep(2.0)   # dar tiempo al GC de Chrome
                try:
                    # Intentar recuperar la pestaña navegando a about:blank primero
                    await page.goto("about:blank", timeout=5000)
                    await asyncio.sleep(0.5)
                except Exception:
                    pass
            else:
                self.warn(f"Error [{url[:50]}]: {err[:60]}")
            return False

    async def autonomous_navigation(self, page):
        """
        Navega autónomamente durante nav_duration segundos.
        Estrategia: extraer links, visitarlos en profundidad, scroll en cada uno.
        """
        self.step("🤖", f"Iniciando navegación autónoma ({self.nav_duration//60} min)...",
                  style="bold magenta")

        end_time   = time.time() + self.nav_duration
        queue      = []
        pages_done = 0

        # Extraer datos de la página actual primero
        data = await self.inject_and_extract(page)
        if data:
            self.pages_data.append(data)
            pages_done += 1

        # Obtener links de la home
        queue = await self.get_internal_links(page)
        random.shuffle(queue)

        with Progress(
            SpinnerColumn(),
            TextColumn("[bold cyan]{task.description}"),
            BarColumn(bar_width=30),
            TextColumn("[green]{task.completed}/{task.total} págs"),
            TimeElapsedColumn(),
            console=console,
            transient=False
        ) as progress:
            task = progress.add_task(
                f"Navegando {self.domain}...",
                total=50   # objetivo máximo por sesión
            )

            while time.time() < end_time and queue:
                url = queue.pop(0)
                if url in self.visited_urls:
                    continue

                remaining = int(end_time - time.time())
                progress.update(task, description=f"[{remaining}s restantes] {url[-50:]}")

                ok = await self.navigate_page(page, url)
                if ok:
                    pages_done += 1
                    progress.advance(task)
                    # Obtener nuevos links de la página actual
                    new_links = await self.get_internal_links(page)
                    for lnk in new_links:
                        if lnk not in self.visited_urls and lnk not in queue:
                            queue.append(lnk)
                    random.shuffle(queue[:10])

                # Pausa entre páginas — equilibrio velocidad/estabilidad
                if time.time() < end_time:
                    await asyncio.sleep(random.uniform(0.8, 1.4))

                # Cada 10 páginas: pausa de limpieza para que Chrome libere memoria
                if pages_done > 0 and pages_done % 10 == 0:
                    progress.update(task, description=f"[{int(end_time-time.time())}s] Limpiando memoria de Chrome...")
                    await page.goto("about:blank", timeout=5000)
                    await asyncio.sleep(1.5)
                    progress.update(task, description=f"[{int(end_time-time.time())}s] Reanudando navegación...")

            # Si el queue se vacía antes de tiempo, intentar extraer más links de la última página
            if time.time() < end_time:
                remaining = int(end_time - time.time())
                extra_links = await self.get_internal_links(page)
                for lnk in extra_links:
                    if lnk not in self.visited_urls:
                        queue.append(lnk)
                if queue:
                    progress.update(task, description=f"[{remaining}s] Continuando con links adicionales...")
                    # Seguir el loop principal con los nuevos links
                    while time.time() < end_time and queue:
                        url = queue.pop(0)
                        if url in self.visited_urls:
                            continue
                        remaining = int(end_time - time.time())
                        progress.update(task, description=f"[{remaining}s restantes] {url[-50:]}")
                        ok = await self.navigate_page(page, url)
                        if ok:
                            pages_done += 1
                            progress.advance(task)
                            new_links = await self.get_internal_links(page)
                            for lnk in new_links:
                                if lnk not in self.visited_urls and lnk not in queue:
                                    queue.append(lnk)
                        if time.time() < end_time:
                            await asyncio.sleep(random.uniform(0.8, 1.4))

        self.success(f"Navegación completada — {pages_done} páginas visitadas, "
                     f"{len(self.visited_urls)} URLs únicas")

    async def wait_for_cdp(self, timeout=15) -> bool:
        import urllib.request
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                urllib.request.urlopen(f"http://localhost:{self.cdp_port}/json/version", timeout=2)
                return True
            except Exception:
                await asyncio.sleep(0.4)
        return False

    def build_scan_data(self) -> dict:
        """Agrega todos los datos de páginas en un resumen."""
        tools = []
        all_s = {}
        all_gtm_containers = []

        for pg in self.pages_data:
            adobe  = pg.get("adobe", {})
            google = pg.get("google", {})
            if adobe.get("s_object"):     tools.append("Adobe Analytics (AppMeasurement / s object)")
            if adobe.get("web_sdk"):      tools.append("Adobe Web SDK (Alloy)")
            if adobe.get("satellite"):    tools.append("Adobe Launch / Tags (_satellite)")
            if adobe.get("digitalData"):  tools.append("W3C Digital Data Layer (digitalData)")
            if adobe.get("acdl"):         tools.append("Adobe Client Data Layer (ACDL)")
            if google.get("dataLayer"):   tools.append("Google Tag Manager (dataLayer)")
            if google.get("gtag"):        tools.append("Google Tag (gtag / GA4)")
            if google.get("ua_trackers"): tools.append("Universal Analytics (ga)")
            teal = pg.get("tealium", {})
            if teal.get("detected"):      tools.append("Tealium iQ")
            if teal.get("data_layer") or teal.get("utag_data"): tools.append("Tealium Data Layer (utag.data)")
            if pg.get("segment", {}).get("detected"): tools.append("Segment")
            if pg.get("piano", {}).get("detected"):   tools.append("Piano Analytics / AT Internet")
            if pg.get("commanders_act", {}).get("detected"): tools.append("Commanders Act (TagCommander)")
            if pg.get("matomo", {}).get("detected"):  tools.append("Matomo")
            all_s.update(adobe.get("s_object", {}))
            all_gtm_containers.extend(pg.get("gtm_containers", []))

        for tool in self.network_hits:
            tools.append(f"{tool} [network beacon]")

        first = self.pages_data[0] if self.pages_data else {}

        return {
            "url":            self.url,
            "domain":         self.domain,
            "title":          first.get("title",""),
            "scan_time":      datetime.now().isoformat(),
            "pages_visited":  len(self.pages_data),
            "duration_s":     round(time.time()-self.start_time, 1) if self.start_time else 0,
            "tools_detected": list(dict.fromkeys(tools)),   # dedup preservando orden
            "report_suite":   first.get("adobe",{}).get("report_suite",""),
            "gtm_containers": list(set(all_gtm_containers)),
        }

    def print_summary(self, scan_data):
        console.print()
        console.rule("[bold green]📊 RESUMEN — WORMLYTICS[/bold green]")
        console.print()

        t = Table(box=box.ROUNDED, border_style="green", show_header=False, padding=(0,2))
        t.add_column("k", style="bold", min_width=25)
        t.add_column("v", style="cyan")
        t.add_row("Dominio",           scan_data["domain"])
        t.add_row("Páginas visitadas", str(scan_data["pages_visited"]))
        t.add_row("Duración",          f"{scan_data['duration_s']}s")
        t.add_row("Herramientas",      str(len(scan_data["tools_detected"])))
        t.add_row("Variables AA",      str(sum(len(p.get("adobe",{}).get("s_object",{})) for p in self.pages_data)))
        t.add_row("Eventos GTM",       str(sum(len(p.get("google",{}).get("dataLayer",[])) for p in self.pages_data)))
        t.add_row("Eventos RT",        str(sum(len(p.get("runtime_events",[])) for p in self.pages_data)))
        t.add_row("Beacons de red",    str(sum(len(v) for v in self.network_hits.values())))
        console.print(t)
        console.print()

        for tool in scan_data["tools_detected"]:
            console.print(f"  [bold green]✓[/bold green]  {tool}")
        console.print()

    async def run(self):
        self.start_time = time.time()
        console.print(BANNER)
        console.print()

        chrome_path = get_chrome_path()
        profile_dir = str(Path.home() / ".wormlytics_profile")

        # Determinar nombre de output
        if not self.output_file:
            safe_domain = re.sub(r"[^\w\-]", "_", self.domain)
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            self.output_file = f"wormlytics_{safe_domain}_{ts}.xlsx"

        console.print(Panel(
            f"[bold]URL:[/bold]           [cyan]{self.url}[/cyan]\n"
            f"[bold]Scope:[/bold]         {self.scope_path or '(todo el dominio)'}\n"
            f"[bold]Navegador:[/bold]     {chrome_path or 'Playwright Chromium (fallback)'}\n"
            f"[bold]Modo:[/bold]          {'Manual — tú aceptas cookies, luego script navega' if self.manual_mode else 'Automático'}\n"
            f"[bold]Navegación:[/bold]    {self.nav_duration}s ({self.nav_duration//60} min {self.nav_duration%60}s)\n"
            f"[bold]Output Excel:[/bold]  [yellow]{self.output_file}[/yellow]",
            title="[bold green]⚙  Configuración[/bold green]", border_style="green"
        ))
        console.print()

        async with async_playwright() as p:

            if chrome_path:
                # ── Lanzar Chrome real con CDP ─────────────────────────────
                self.step("🚀", "Lanzando Chrome real con puerto CDP...", f":{self.cdp_port}")
                Path(profile_dir).mkdir(parents=True, exist_ok=True)

                cmd = [chrome_path,
                       f"--remote-debugging-port={self.cdp_port}",
                       f"--user-data-dir={profile_dir}",
                       "--no-first-run", "--no-default-browser-check",
                       self.url]
                proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

                self.step("⏳", "Esperando que Chrome arranque...")
                if not await self.wait_for_cdp():
                    self.error(f"Chrome no respondió en puerto {self.cdp_port}")
                    self.warn("¿Tienes otro Chrome abierto? Ciérralo e intenta de nuevo.")
                    proc.terminate()
                    sys.exit(1)
                self.success(f"Chrome listo en puerto {self.cdp_port}")

                if self.manual_mode:
                    # ── Modo manual: usuario acepta cookies ────────────────
                    console.print()
                    console.print(Panel(
                        f"[bold white]Chrome está abierto en:[/bold white] [cyan]{self.url}[/cyan]\n\n"
                        "  [bold yellow]👉 Tu turno:[/bold yellow]\n\n"
                        "     1. Acepta el banner de cookies\n"
                        "     2. Haz alguna navegación si quieres\n"
                        "     3. Cuando estés listo → [bold green]ENTER[/bold green]\n\n"
                        "  [dim]Después el script navegará solo durante "
                        f"{self.nav_duration//60} minutos y generará el Excel.[/dim]",
                        title="[bold green]🙋 TU TURNO[/bold green]",
                        border_style="bright_green", padding=(1,3)
                    ))
                    console.print()
                    input("  ⏎  Pulsa ENTER cuando hayas aceptado las cookies...")
                    console.print()

                # Conectar via CDP
                self.step("🔌", "Conectando via CDP...")
                try:
                    browser = await p.chromium.connect_over_cdp(f"http://localhost:{self.cdp_port}")
                    self.success("Conectado al Chrome")
                except Exception as e:
                    self.error(f"Error CDP: {e}")
                    proc.terminate()
                    sys.exit(1)

                ctx  = browser.contexts[0] if browser.contexts else None
                if not ctx or not ctx.pages:
                    self.error("No se encontró contexto o pestañas")
                    proc.terminate()
                    sys.exit(1)

                # Seleccionar pestaña con la URL objetivo
                page = next((pg for pg in ctx.pages if self.domain in pg.url), ctx.pages[-1])
                self.success(f"Pestaña activa: {page.url[:80]}")

            else:
                # ── Fallback: Playwright Chromium ─────────────────────────
                self.warn("Chrome real no encontrado — usando Playwright Chromium")
                browser = await p.chromium.launch(
                    headless=False,
                    args=["--disable-blink-features=AutomationControlled","--no-sandbox"]
                )
                ctx  = await browser.new_context(viewport={"width":1440,"height":900})
                # Registrar interceptores como init script — pre-load en cada navegación
                await ctx.add_init_script(JS_INJECT)
                self.success("Interceptores registrados como init script")
                page = await ctx.new_page()

                if self.manual_mode:
                    self.step("🌐", "Abriendo URL...", self.url)
                    await page.goto(self.url, wait_until="domcontentloaded", timeout=30000)
                    await page.wait_for_timeout(2000)

                    console.print()
                    console.print(Panel(
                        "  [bold yellow]👉 Tu turno:[/bold yellow]\n\n"
                        "     1. Acepta el banner de cookies\n"
                        "     2. Pulsa [bold green]ENTER[/bold green] cuando estés listo",
                        title="[bold green]🙋 TU TURNO[/bold green]",
                        border_style="bright_green", padding=(1,2)
                    ))
                    input("  ⏎  ENTER cuando estés listo...")
                    console.print()
                else:
                    self.step("🌐", "Navegando a la URL inicial...", self.url)
                    await page.goto(self.url, wait_until="domcontentloaded", timeout=30000)
                    await page.wait_for_timeout(3000)

                proc = None

            # ── Interceptor de red ─────────────────────────────────────────
            self.step("🕸️", "Activando interceptor de red...")
            page.on("request", self._on_request)

            # ── Navegación autónoma ────────────────────────────────────────
            await self.autonomous_navigation(page)

            # ── Agregado y resumen ─────────────────────────────────────────
            self.step("🧠", "Agregando datos de todas las páginas...")
            scan_data = self.build_scan_data()
            self.print_summary(scan_data)

            # ── Exportar Excel ─────────────────────────────────────────────
            if not XLSX_OK:
                self.error("openpyxl no instalado — pip install openpyxl")
            else:
                self.step("📊", "Generando Excel...", self.output_file)
                output_path = str(Path(self.output_file).resolve())
                clean_pages = _sanitize_for_json(self.pages_data)
                build_excel(scan_data, clean_pages, self.network_hits, output_path)
                self.success(f"Excel guardado: {output_path}")

                # También guardar JSON raw por si acaso
                json_path = output_path.replace(".xlsx", "_raw.json")
                raw = _sanitize_for_json({
                    "meta": scan_data,
                    "pages": self.pages_data,
                    "network_hits": self.network_hits,
                    "wormlytics_version": "5.5"
                })
                Path(json_path).write_text(
                    json.dumps(raw, indent=2, ensure_ascii=False, default=_safe_json_default),
                    encoding="utf-8"
                )
                self.success(f"JSON raw guardado: {json_path}")

            total = round(time.time()-self.start_time, 1)
            console.print()
            console.print(Panel(
                f"[bold green]✓ Completado en {total}s  "
                f"({len(self.pages_data)} páginas  |  {self.step_count} pasos)[/bold green]\n\n"
                f"[bold yellow]📂 Excel:[/bold yellow] [cyan]{self.output_file}[/cyan]\n"
                "[dim]Pégalo en ChatGPT / Claude para generar tu SDR personalizado.[/dim]",
                border_style="green", padding=(1,3)
            ))

            if chrome_path and proc:
                console.print("  [dim]Chrome sigue abierto. Ciérralo manualmente cuando quieras.[/dim]")
                try:
                    await browser.close()
                except Exception:
                    pass

# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────




def main():
    parser = argparse.ArgumentParser(
        description="Wormlytics v4.0 — Analytics Data Layer Hunter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  # Modo manual (tú aceptas cookies, luego el bot navega 5 min)
  python wormlytics.py https://www.example.com --manual

  # Modo automático total (5 min de navegación desde el inicio)
  python wormlytics.py https://www.example.com

  # Especificar duración y output
  python wormlytics.py https://www.example.com --duration 180 --output mi_analisis.xlsx

  # Modo manual con duración corta para pruebas
  python wormlytics.py https://www.example.com --manual --duration 60
        """
    )
    parser.add_argument("url",            help="URL a analizar (incluir https://)")
    parser.add_argument("--manual","-m",  action="store_true",
                                          help="Modo manual: tú navegas y aceptas cookies primero, luego el bot toma el control")
    parser.add_argument("--duration","-d",type=int, default=NAV_DURATION, metavar="S",
                                          help=f"Duración de la navegación autónoma en segundos (default: {NAV_DURATION})")
    parser.add_argument("--output","-o",  metavar="FILE",
                                          help="Nombre del archivo Excel de salida (default: wormlytics_<domain>_<timestamp>.xlsx)")
    parser.add_argument("--port",         type=int, default=CDP_PORT, metavar="N",
                                          help=f"Puerto CDP para Chrome (default: {CDP_PORT})")
    parser.add_argument("--scope",        metavar="PATH",
                                          help="Restringir navegación a URLs que contengan este path "
                                               "(ej: --scope /entradas  navega solo dentro de /entradas)")

    args = parser.parse_args()

    if not args.url.startswith(("http://","https://")):
        console.print("[bold red]Error: La URL debe incluir http:// o https://[/bold red]")
        sys.exit(1)

    if not XLSX_OK:
        console.print("[bold red]Error: openpyxl no instalado. Ejecuta: pip install openpyxl[/bold red]")
        sys.exit(1)

    bot = Wormlytics(
        url=args.url,
        output_file=args.output,
        cdp_port=args.port,
        nav_duration=args.duration,
        manual_mode=args.manual,
        scope_path=args.scope,
    )
    asyncio.run(bot.run())

if __name__ == "__main__":
    main()
