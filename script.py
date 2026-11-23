#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bot de trading por Telegram para Binance USDⓈ-M Futures en Hedge Mode.

- Una sola aplicación Python; datos persistentes en trading_bot_data.json.
- Lee de señales solo símbolo, BUY/SELL y rango; ignora TP/SL/leverage del mensaje.
- Conserva perfiles y extensiones de rango del Excel del usuario.
- BTC/ETH: leverage fijo 100x, Stop Loss editable (por defecto -300% ROI).
- Altcoins: leverage seleccionable 20x-30x y Stop Loss editable 80%-120% ROI.
- Stop de todos los perfiles desde el promedio ponderado: ROI / leverage / 100.
- Coloca entradas LIMIT; tras cada fill coloca LIMIT parcial y mantiene un
  STOP_MARKET de protección de la pata más un TRAILING_STOP_MARKET agregado.
- DEMO por defecto; LIVE requiere ALLOW_LIVE_TRADING=YES y confirmación por Telegram.
- HEDGE: usa positionSide LONG/SHORT en entradas y salidas.

Requiere Python 3.10+ y python-telegram-bot 22.8 + requests.
"""
from __future__ import annotations

import asyncio
import copy
import hashlib
import hmac
import json
import logging
import math
import os
import re
import time
import urllib.parse
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP
from pathlib import Path
from typing import Any

import requests
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "trading_bot_data.json"
ENV_FILE = APP_DIR / ".env"
# Binance USDⓈ-M Futures: DEMO por defecto; LIVE requiere doble habilitación.
BINANCE_URLS = {
    "DEMO": "https://demo-fapi.binance.com",
    "LIVE": "https://fapi.binance.com",
}
RECV_WINDOW = 5000
POLL_SECONDS = 3
MAX_ENTRIES = 20
LOG = logging.getLogger("trading_bot")
SESSION = requests.Session()

# ----------------------------- Persistencia ---------------------------------
DEFAULT_DATA: dict[str, Any] = {
    "version": 9, "owner_user_id": None, "setup": None,
    "wizard": {}, "pending": {}, "live_confirm": {}, "last_signal": {}, "trades": {},
}

def load_env() -> None:
    if not ENV_FILE.exists():
        return
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if not s or s.startswith("#") or "=" not in s:
            continue
        k, v = s.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

def merge_defaults(obj: Any, defaults: Any) -> Any:
    if isinstance(defaults, dict):
        out = copy.deepcopy(obj) if isinstance(obj, dict) else {}
        for k, v in defaults.items():
            out[k] = merge_defaults(out[k], v) if k in out else copy.deepcopy(v)
        return out
    return obj

def save_data() -> None:
    tmp = DATA_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(DATA, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(DATA_FILE)

def load_data() -> dict[str, Any]:
    if not DATA_FILE.exists():
        result = copy.deepcopy(DEFAULT_DATA)
        tmp = DATA_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        tmp.replace(DATA_FILE)
        return result
    try:
        raw = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        old_version = int(raw.get("version", 0) or 0)
        # v5 corrige las plantillas BUY/SELL separadas del Excel.
        if old_version < 5 and isinstance(raw.get("setup"), dict):
            raw["setup"]["btc_allocations"] = None
            raw["setup"]["alt_allocations"] = None
        # v8 fija BTC/ETH a 100x y el bot a Hedge Mode.
        # v9 añade Stop Loss editable por perfil y restringe altcoins a 20-30x.
        if isinstance(raw.get("setup"), dict):
            setup = raw["setup"]
            setup["btc_leverage"] = 100
            setup.setdefault("btc_stop_roi_pct", 300.0)
            setup.setdefault("alt_stop_roi_pct", 120.0)
            try:
                old_alt_lev = int(float(setup.get("alt_leverage") or 20))
            except (TypeError, ValueError):
                old_alt_lev = 20
            setup["alt_leverage"] = old_alt_lev if 20 <= old_alt_lev <= 30 else 20
        if old_version < 8:
            for uid, wizard in list(raw.get("wizard", {}).items()):
                if wizard.get("kind") == "setup":
                    raw["wizard"].pop(uid, None)
            for trade in raw.get("trades", {}).values():
                if trade.get("status") in {"ACTIVE", "SUBMITTING"} and trade.get("position_side") not in {"LONG", "SHORT"}:
                    trade["status"] = "LEGACY_REVIEW"
                    trade["legacy_review_warning_sent"] = False
        if old_version < 9:
            # Los asistentes de configuración antiguos tenían otro número de pasos.
            for uid, wizard in list(raw.get("wizard", {}).items()):
                if wizard.get("kind") == "setup":
                    raw["wizard"].pop(uid, None)
        raw["version"] = 9
        return merge_defaults(raw, DEFAULT_DATA)
    except Exception as exc:
        backup = DATA_FILE.with_name("trading_bot_data.corrupto.json")
        try:
            DATA_FILE.replace(backup)
        except OSError:
            pass
        LOG.error("No se pudo leer el estado; respaldo en %s: %s", backup, exc)
        return copy.deepcopy(DEFAULT_DATA)

load_env()
DATA = load_data()

# ------------------------------- Utilidades ---------------------------------
def parse_num(value: str) -> float:
    s = value.strip().replace(" ", "")
    if not s:
        raise ValueError("Valor vacío.")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        # En señales, 0,1234 es decimal; 1,234 se interpreta como miles.
        if re.fullmatch(r"\d{1,3},\d{3}", s) and not s.startswith("0,"):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    n = float(s)
    if not math.isfinite(n):
        raise ValueError("Número no válido.")
    return n

def excel_round(x: float, digits: int) -> float:
    return float(Decimal(str(x)).quantize(Decimal("1").scaleb(-digits), rounding=ROUND_HALF_UP))

def fmt(x: float | None, digits: int = 8) -> str:
    if x is None:
        return "—"
    s = f"{x:,.{digits}f}"
    return s.rstrip("0").rstrip(".") if "." in s else s

def api_num(x: float, digits: int = 12) -> str:
    """Formato decimal sin separadores de miles para parámetros de Binance."""
    s = f"{float(x):.{digits}f}".rstrip("0").rstrip(".")
    return s or "0"

def pct_allocations(raw: str, count: int, defaults: list[float]) -> list[float]:
    if raw.strip().upper() in {"DEFAULT", "PREDETERMINADO", "EXCEL"}:
        return defaults[:]
    parts = [p for p in re.split(r"[;,\s]+", raw.strip()) if p]
    vals = [parse_num(p) for p in parts]
    if len(vals) != count or any(v < 0 or v > 100 for v in vals) or abs(sum(vals)-100) > 0.01:
        raise ValueError(f"Ingresa {count} porcentajes que sumen 100, o escribe DEFAULT.")
    vals[-1] += 100 - sum(vals)
    return vals

def floor_step(value: float, step: str) -> float:
    d, st = Decimal(str(value)), Decimal(str(step))
    if st <= 0:
        raise ValueError("stepSize inválido del exchange.")
    return float((d / st).to_integral_value(rounding=ROUND_DOWN) * st)

def quantize_price(value: float, tick: str, direction: str = "nearest") -> float:
    d, t = Decimal(str(value)), Decimal(str(tick))
    mode = ROUND_HALF_UP if direction == "nearest" else ROUND_UP if direction == "up" else ROUND_DOWN
    return float((d / t).to_integral_value(rounding=mode) * t)

def decimal_places(step: str) -> int:
    return max(0, -Decimal(str(step)).normalize().as_tuple().exponent)

# ---------------------- Perfiles del Excel original -------------------------
BTC_ALLOC_BUY = {
    3: [25, 30, 45],
    4: [15, 20, 25, 40],
    5: [10, 15, 20, 25, 30],
    6: [11, 11.5, 12.5, 14, 18, 33],
}
BTC_ALLOC_SELL = {
    3: [25, 30, 45],
    4: [15, 20, 30, 35],
    5: [12, 12, 13, 18, 45],
    6: [11, 11.5, 12.5, 14, 18, 33],
}
ALT_ALLOC_BUY = {
    3: [25, 30, 45],
    4: [15, 20, 30, 35],
    6: [5, 8, 12, 15, 22, 38],
}
ALT_ALLOC_SELL = {
    3: [25, 30, 45],
    4: [15, 20, 30, 35],
    6: [10, 10, 10, 20, 20, 30],
}
ALT_DECIMALS = {
    "ADA": 0,
    "XRP": 1, "EOS": 1, "QTUM": 1, "IOTA": 1,
    "ETC": 2, "LINK": 2,
    "LTC": 3, "DASH": 3, "ZEC": 3, "XMR": 3,
}
ALT_PERCENT_BY_DEC = {0: 10.0, 1: 10.0, 2: 10.0, 3: 4.845}


def asset_class(symbol: str) -> str:
    base = symbol.upper().replace("/", "").removesuffix("USDT")
    return "btc_eth" if base in {"BTC", "ETH"} else "altcoin"


def default_count_percent_leverage(kind: str, count: int, decimals: int = 3,
                                   direction: str = "BUY") -> tuple[float, float]:
    if kind == "btc_eth":
        # El usuario fijó BTC/ETH a 100x. Se conservan los porcentajes de
        # capital del Excel, pero el apalancamiento predeterminado es 100x.
        if count == 3: return 8.333, 100
        if count == 4: return 3.333, 100
        if count == 5: return 2.5, 100
        if count == 6: return (3.333 if direction == "BUY" else 2.5), 100
        return 3.333, 100
    if count == 4:
        return (3.333, 20) if direction == "BUY" else (8.333, 20)
    if count == 6:
        return 8.333, 20
    if direction == "SELL":
        return 10.0, 20
    return ALT_PERCENT_BY_DEC.get(decimals, 4.845), 20


def default_weights(kind: str, count: int, direction: str = "BUY", decimals: int = 3) -> list[float]:
    if kind == "btc_eth":
        table = BTC_ALLOC_BUY if direction == "BUY" else BTC_ALLOC_SELL
    else:
        table = ALT_ALLOC_BUY if direction == "BUY" else ALT_ALLOC_SELL
    if kind == "altcoin" and count == 3 and direction == "SELL" and decimals == 2:
        return [30, 30, 40]
    if count in table:
        return list(table[count])
    return [100 / count] * (count - 1) + [100 - (100 / count) * (count - 1)]


def resolve_profile(symbol: str, setup: dict[str, Any], direction: str = "BUY") -> dict[str, Any]:
    base = symbol.upper().replace("/", "").removesuffix("USDT")
    kind = asset_class(symbol)
    direction = "SELL" if direction == "SELL" else "BUY"
    count = int(setup["btc_count"] if kind == "btc_eth" else setup["alt_count"])
    decimals = 3 if kind == "btc_eth" else ALT_DECIMALS.get(base, 3)
    default_pct, default_lev = default_count_percent_leverage(kind, count, decimals, direction)
    pct_override = setup.get("btc_operate") if kind == "btc_eth" else setup.get("alt_operate")
    lev_override = setup.get("alt_leverage") if kind == "altcoin" else None
    operate = float(pct_override) if pct_override is not None else default_pct
    # BTC/ETH permanecen en 100x; altcoins permiten elegir entre 20x y 30x.
    leverage = 100.0 if kind == "btc_eth" else (float(lev_override) if lev_override is not None else 20.0)
    if kind == "altcoin" and (not leverage.is_integer() or not 20 <= leverage <= 30):
        raise ValueError("El apalancamiento de altcoins debe ser entero y estar entre 20x y 30x.")
    alloc_setting = setup.get("btc_allocations") if kind == "btc_eth" else setup.get("alt_allocations")
    allocs = list(alloc_setting) if alloc_setting else default_weights(kind, count, direction, decimals)
    # En las hojas de 3 entradas la cantidad sigue la clasificación de MONEDAS;
    # las plantillas genéricas de 4 y 6 entradas redondean cantidad a 3 decimales.
    qdec = 3 if kind == "btc_eth" or (kind == "altcoin" and count in {4, 6}) else decimals
    pdec = 2 if kind == "btc_eth" else {0: 5, 1: 4, 2: 3, 3: 2}.get(decimals, 2)
    # Las plantillas genéricas de 4/6 entradas y BTC/ETH 3 tienen
    # precisiones de precio propias, también distintas entre algunas hojas.
    if kind == "btc_eth" and count == 3 and direction == "SELL":
        pdec = 4
    elif kind == "altcoin" and count == 4:
        pdec = 4 if direction == "BUY" else 2
    elif kind == "altcoin" and count == 6:
        pdec = 4
    # Porcentajes de rango y distribuciones direccionales tomados de cada hoja.
    if kind == "btc_eth" and count == 5:
        range_fraction, sheet = (0.05 if direction == "BUY" else 0.0485), "BTC ETH 5 ENT."
    elif kind == "btc_eth" and count == 6:
        range_fraction, sheet = 0.12, "BTC ETH 6 ENT."
    elif kind == "btc_eth":
        range_fraction, sheet = None, f"BTC ETH {count} ENT."
    elif count == 3:
        range_fraction, sheet = (0.03 if direction == "BUY" else 0.015), f"{decimals} DEC."
    elif count == 4:
        range_fraction, sheet = None, "4 ENT."
    elif count == 6:
        range_fraction, sheet = None, "6 ENT."
    else:
        range_fraction, sheet = None, f"ALTCOIN {count} ENT."
    # Stop Loss editable por perfil, aplicado igual que en el Excel:
    # movimiento adverso del precio = ROI de riesgo / leverage / 100.
    if kind == "btc_eth":
        stop_roi_pct = float(setup.get("btc_stop_roi_pct", 300.0))
        if not 1.0 <= stop_roi_pct <= 1000.0:
            raise ValueError("El Stop Loss BTC/ETH debe estar entre 1% y 1000% ROI.")
    else:
        stop_roi_pct = float(setup.get("alt_stop_roi_pct", 120.0))
        if not 80.0 <= stop_roi_pct <= 120.0:
            raise ValueError("El Stop Loss de altcoins debe estar entre 80% y 120% ROI.")
    return {
        "kind": kind, "label": "BTC / ETH" if kind == "btc_eth" else "ALTCOIN",
        "sheet": sheet, "count": count, "allocations": allocs,
        "operate_pct": operate, "leverage": leverage,
        "quantity_decimals": qdec, "price_decimals": pdec,
        "range_fraction": range_fraction, "stop_roi_pct": stop_roi_pct,
        "margin_type": "CROSSED" if kind == "btc_eth" or (kind == "altcoin" and count in {4, 6}) else "ISOLATED",
        "direction": direction,
    }

def build_plan(signal: dict[str, Any], setup: dict[str, Any]) -> dict[str, Any]:
    profile = resolve_profile(signal["symbol"], setup, signal["direction"])
    fraction = profile["range_fraction"]
    if fraction is None:
        fraction = (signal["max_price"] - signal["min_price"]) / signal["min_price"]
    if signal["direction"] == "BUY":
        first, last = signal["max_price"], signal["max_price"] * (1 - fraction)
    else:
        first, last = signal["min_price"], signal["min_price"] * (1 + fraction)
    n = profile["count"]
    prices = [first] if n == 1 else [first + (last-first)*i/(n-1) for i in range(n)]
    capital = float(setup["capital"])
    total_margin = capital * profile["operate_pct"] / 100
    entries = []
    for i, (weight, price) in enumerate(zip(profile["allocations"], prices), 1):
        p = excel_round(price, profile["price_decimals"])
        margin = total_margin * weight / 100
        nominal = margin * profile["leverage"]
        qty = excel_round(nominal / p, profile["quantity_decimals"])
        entries.append({"number": i, "weight": weight, "price": p, "margin": margin,
                        "notional": nominal, "quantity_excel": qty})
    total_qty = sum(x["quantity_excel"] for x in entries)
    notional = sum(x["quantity_excel"] * x["price"] for x in entries)
    avg = notional / total_qty if total_qty else 0
    return {"profile": profile, "entries": entries, "total_margin": total_margin,
            "total_notional": sum(e["notional"] for e in entries),
            "average_price": avg, "total_qty_excel": total_qty, "range_fraction": fraction}

# ------------------------- Lector de señales --------------------------------
def parse_signal(text: str) -> dict[str, Any]:
    if not text:
        raise ValueError("El mensaje está vacío.")
    raw = text.strip()
    match = re.search(r"#?\b([A-Z0-9]{2,20})\s*/?\s*USDT\b", raw, re.I)
    if not match:
        raise ValueError("No encontré el par USDT (ejemplo: #BTC/USDT o #SOLUSDT).")
    base = match.group(1).upper()
    # La dirección se lee del encabezado de la señal, antes de las filas TP/SL.
    header_lines = []
    for line in raw.splitlines():
        if re.match(r"^\s*(?:TP\s*\d*|TAKE\s*PROFIT|TARGET|STOP\s*LOSS|STOP\s*\d*|SL)\b", line, re.I):
            break
        header_lines.append(line)
    header = "\n".join(header_lines) or raw
    dirs = list(re.finditer(r"\b(BUY|LONG|COMPRA|COMPRAR|SELL|SHORT|VENTA|VENDER)\b", header, re.I))
    if not dirs:
        raise ValueError("No encontré BUY/LONG o SELL/SHORT en el encabezado de la señal.")
    direction = "BUY" if dirs[0].group(1).upper() in {"BUY", "LONG", "COMPRA", "COMPRAR"} else "SELL"
    rm = re.search(r"(?:Rango(?:\s+de\s+entrada)?|Entry\s*Range)\s*:\s*([0-9][0-9.,]*)\s*[-–—]\s*([0-9][0-9.,]*)", raw, re.I)
    if not rm:
        raise ValueError("No encontré el rango. Ejemplo: Rango: 0.1690 - 0.1890.")
    a, b = parse_num(rm.group(1)), parse_num(rm.group(2))
    if a <= 0 or b <= 0:
        raise ValueError("Los precios deben ser mayores que cero.")
    return {"symbol": f"{base}USDT", "display_symbol": f"{base}/USDT", "direction": direction,
            "min_price": min(a, b), "max_price": max(a, b)}

# ------------------------- Binance USDⓈ-M Futures ---------------------------
def configured_binance_mode() -> str:
    mode = os.getenv("BINANCE_ENV", "DEMO").strip().upper()
    if mode not in BINANCE_URLS:
        raise RuntimeError("BINANCE_ENV debe ser DEMO o LIVE.")
    return mode

def live_trading_allowed() -> bool:
    return os.getenv("ALLOW_LIVE_TRADING", "NO").strip().upper() in {"YES", "SI", "TRUE", "1"}

def mode_label(mode: str | None = None) -> str:
    selected = mode or configured_binance_mode()
    return "PRODUCCIÓN (DINERO REAL)" if selected == "LIVE" else "DEMO (SIMULACIÓN)"

class BinanceUSDM:
    def __init__(self, mode: str | None = None) -> None:
        self.mode = (mode or configured_binance_mode()).upper()
        if self.mode not in BINANCE_URLS:
            raise RuntimeError("Entorno Binance inválido: usa DEMO o LIVE.")
        self.base = BINANCE_URLS[self.mode]
        # Credenciales separadas para poder seguir supervisando operaciones DEMO
        # aunque el usuario cambie el entorno de nuevas señales a LIVE.
        self.api_key = os.getenv(f"BINANCE_{self.mode}_API_KEY", os.getenv("BINANCE_API_KEY", "")).strip()
        self.secret = os.getenv(f"BINANCE_{self.mode}_API_SECRET", os.getenv("BINANCE_API_SECRET", "")).strip()

    def credentials_ok(self) -> bool:
        return bool(self.api_key and self.secret)

    def request(self, method: str, path: str, params: dict[str, Any] | None = None,
                signed: bool = False) -> Any:
        p = {k: v for k, v in (params or {}).items() if v is not None}
        headers = {"X-MBX-APIKEY": self.api_key} if self.api_key else {}
        if self.mode == "LIVE" and not live_trading_allowed():
            raise RuntimeError("Modo LIVE bloqueado. En .env debes poner ALLOW_LIVE_TRADING=YES y confirmar cada operación en Telegram.")
        if signed:
            if not self.credentials_ok():
                raise RuntimeError("Faltan BINANCE_API_KEY y BINANCE_API_SECRET en .env.")
            p["timestamp"] = int(time.time() * 1000)
            p["recvWindow"] = RECV_WINDOW
            query = urllib.parse.urlencode(p, doseq=True)
            p["signature"] = hmac.new(self.secret.encode(), query.encode(), hashlib.sha256).hexdigest()
        url = self.base + path
        try:
            r = SESSION.request(method.upper(), url, params=p if method.upper() in {"GET", "DELETE"} else None,
                                data=p if method.upper() not in {"GET", "DELETE"} else None,
                                headers=headers, timeout=15)
        except requests.RequestException as exc:
            raise RuntimeError(f"Error de red Binance {self.mode}: {exc}") from exc
        try:
            body = r.json()
        except ValueError:
            body = {"msg": r.text[:500]}
        if not r.ok:
            code = body.get("code") if isinstance(body, dict) else "?"
            msg = body.get("msg", body) if isinstance(body, dict) else body
            raise RuntimeError(f"Binance {r.status_code} [{code}]: {msg}")
        return body

    def exchange_info(self) -> dict[str, Any]:
        return self.request("GET", "/fapi/v1/exchangeInfo")

    def symbol_rules(self, symbol: str) -> dict[str, Any]:
        info = self.exchange_info()
        item = next((s for s in info["symbols"] if s["symbol"] == symbol), None)
        if not item or item.get("status") != "TRADING" or item.get("contractType") != "PERPETUAL" or item.get("quoteAsset") != "USDT":
            raise ValueError(f"{symbol} no está habilitado como perpetuo USDT en Binance {self.mode}.")
        fs = {f["filterType"]: f for f in item["filters"]}
        lot = fs.get("LOT_SIZE", {})
        price = fs.get("PRICE_FILTER", {})
        notional = fs.get("MIN_NOTIONAL", {})
        return {"symbol": symbol, "step": lot.get("stepSize", "1"), "min_qty": float(lot.get("minQty", "0")),
                "tick": price.get("tickSize", "0.01"), "min_notional": float(notional.get("notional", notional.get("minNotional", 5)))}

    def account(self) -> dict[str, Any]:
        return self.request("GET", "/fapi/v3/account", signed=True)

    def is_hedge_mode(self) -> bool:
        result = self.request("GET", "/fapi/v1/positionSide/dual", signed=True)
        return bool(result.get("dualSidePosition"))

    def ensure_hedge_mode(self) -> None:
        """Activa Hedge Mode si aún no está habilitado.

        Binance puede rechazar el cambio si existe cualquier orden/posición
        abierta en la cuenta, incluso en otro símbolo. En ese caso se detiene
        antes de enviar nuevas órdenes.
        """
        if self.is_hedge_mode():
            return
        try:
            self.request("POST", "/fapi/v1/positionSide/dual",
                         {"dualSidePosition": "true"}, signed=True)
        except RuntimeError as exc:
            if "-4067" in str(exc) or "-4068" in str(exc):
                raise RuntimeError(
                    "Binance no permite activar Hedge Mode porque hay órdenes o posiciones abiertas "
                    "en algún símbolo de la cuenta. Cancela/cierra esas operaciones, vuelve a intentar "
                    "y no se enviaron las entradas de esta señal."
                ) from exc
            raise
        if not self.is_hedge_mode():
            raise RuntimeError("No se pudo confirmar que Binance quedó en Hedge Mode.")

    @staticmethod
    def position_side(direction: str) -> str:
        return "LONG" if direction == "BUY" else "SHORT"

    def positions(self, symbol: str) -> list[dict[str, Any]]:
        rows = self.request("GET", "/fapi/v3/positionRisk", {"symbol": symbol}, signed=True)
        return [p for p in rows if p.get("symbol") == symbol and p.get("positionSide") in {"LONG", "SHORT"}]

    def position(self, symbol: str, direction: str) -> dict[str, Any]:
        side = self.position_side(direction)
        rows = self.positions(symbol)
        return next((p for p in rows if p.get("positionSide") == side),
                    {"symbol": symbol, "positionSide": side, "positionAmt": "0", "entryPrice": "0"})

    def mark_price(self, symbol: str) -> float:
        d = self.request("GET", "/fapi/v1/premiumIndex", {"symbol": symbol})
        return float(d["markPrice"])

    def set_symbol_settings(self, symbol: str, margin_type: str, leverage: int) -> None:
        try:
            self.request("POST", "/fapi/v1/marginType", {"symbol": symbol, "marginType": margin_type}, signed=True)
        except RuntimeError as exc:
            if "-4046" not in str(exc) and "No need to change margin type" not in str(exc):
                raise
        self.request("POST", "/fapi/v1/leverage", {"symbol": symbol, "leverage": leverage}, signed=True)

    def new_limit(self, symbol: str, side: str, qty: float, price: float, client_id: str,
                  position_side: str) -> dict[str, Any]:
        if position_side not in {"LONG", "SHORT"}:
            raise ValueError("Hedge Mode requiere positionSide LONG o SHORT.")
        # Binance no permite enviar reduceOnly en Hedge Mode. La protección
        # se hace especificando siempre la pata LONG/SHORT correcta.
        return self.request("POST", "/fapi/v1/order", {
            "symbol": symbol, "side": side, "type": "LIMIT", "timeInForce": "GTC",
            "quantity": api_num(qty), "price": api_num(price), "positionSide": position_side,
            "newClientOrderId": client_id, "newOrderRespType": "RESULT",
        }, signed=True)

    def new_market_close(self, symbol: str, side: str, qty: float, position_side: str,
                         client_id: str) -> dict[str, Any]:
        if position_side not in {"LONG", "SHORT"}:
            raise ValueError("Hedge Mode requiere positionSide LONG o SHORT.")
        return self.request("POST", "/fapi/v1/order", {
            "symbol": symbol, "side": side, "type": "MARKET", "quantity": api_num(qty),
            "positionSide": position_side, "newClientOrderId": client_id,
            "newOrderRespType": "RESULT",
        }, signed=True)

    def order(self, symbol: str, order_id: int) -> dict[str, Any]:
        return self.request("GET", "/fapi/v1/order", {"symbol": symbol, "orderId": order_id}, signed=True)

    def order_by_client_id(self, symbol: str, client_id: str) -> dict[str, Any]:
        return self.request("GET", "/fapi/v1/order", {"symbol": symbol, "origClientOrderId": client_id}, signed=True)

    def new_trailing(self, symbol: str, side: str, qty: float, activation: float | None,
                     callback: float, client_id: str, position_side: str) -> dict[str, Any]:
        if position_side not in {"LONG", "SHORT"}:
            raise ValueError("Hedge Mode requiere positionSide LONG o SHORT.")
        params: dict[str, Any] = {
            "algoType": "CONDITIONAL", "symbol": symbol, "side": side,
            "type": "TRAILING_STOP_MARKET", "quantity": api_num(qty),
            "positionSide": position_side, "callbackRate": api_num(callback, 2),
            "workingType": "MARK_PRICE", "clientAlgoId": client_id,
        }
        if activation is not None:
            params["activatePrice"] = api_num(activation)
        return self.request("POST", "/fapi/v1/algoOrder", params, signed=True)

    def new_stop_loss(self, symbol: str, side: str, trigger_price: float,
                      client_id: str, position_side: str) -> dict[str, Any]:
        if position_side not in {"LONG", "SHORT"}:
            raise ValueError("Hedge Mode requiere positionSide LONG o SHORT.")
        # Close-all STOP_MARKET: en Hedge Mode solo se permite del lado que
        # cierra la pata (SELL/LONG o BUY/SHORT); sin quantity ni reduceOnly.
        return self.request("POST", "/fapi/v1/algoOrder", {
            "algoType": "CONDITIONAL", "symbol": symbol, "side": side,
            "type": "STOP_MARKET", "positionSide": position_side,
            "closePosition": "true", "triggerPrice": api_num(trigger_price),
            "workingType": "MARK_PRICE", "clientAlgoId": client_id,
        }, signed=True)

    def algo_order(self, algo_id: int) -> dict[str, Any]:
        return self.request("GET", "/fapi/v1/algoOrder", {"algoId": algo_id}, signed=True)

    def algo_order_by_client_id(self, client_id: str) -> dict[str, Any]:
        return self.request("GET", "/fapi/v1/algoOrder", {"clientAlgoId": client_id}, signed=True)

    def cancel_order(self, symbol: str, order_id: int) -> dict[str, Any]:
        return self.request("DELETE", "/fapi/v1/order", {"symbol": symbol, "orderId": order_id}, signed=True)

    def cancel_algo(self, algo_id: int) -> dict[str, Any]:
        return self.request("DELETE", "/fapi/v1/algoOrder", {"algoId": algo_id}, signed=True)

# -------------------------- Salidas por cada fill ----------------------------
def calc_roi_price(entry_price: float, direction: str, leverage: float, roi_pct: float) -> float:
    move = roi_pct / leverage / 100
    return entry_price * (1 + move) if direction == "BUY" else entry_price * (1 - move)


def calc_excel_stop_price(average_entry: float, direction: str, leverage: float,
                          stop_roi_pct: float) -> float:
    """Replica la fórmula del Excel desde el promedio ponderado.

    LONG: avg - ((riskROI / leverage) / 100) * avg
    SHORT: avg + ((riskROI / leverage) / 100) * avg
    """
    if average_entry <= 0 or leverage <= 0 or stop_roi_pct <= 0:
        raise ValueError("No se puede calcular el stop con promedio, leverage o riesgo inválido.")
    adverse_move = stop_roi_pct / leverage / 100
    return average_entry * (1 - adverse_move) if direction == "BUY" else average_entry * (1 + adverse_move)

def exit_price_for_fill(trade: dict[str, Any], entry_price: float, kind: str) -> float:
    setup = trade["setup"]
    prof = trade["profile"]
    rules = trade["rules"]
    if kind == "partial":
        mode, val = setup["tp_mode"], float(setup["tp_value"])
    else:
        mode, val = setup["trail_mode"], float(setup["trail_value"])
    if mode == "PRICE":
        return val
    return calc_roi_price(entry_price, trade["direction"], prof["leverage"], val)

def validate_exit_settings(plan: dict[str, Any], setup: dict[str, Any]) -> None:
    """Evita abrir la posición si los niveles elegidos no están del lado rentable."""
    avg = float(plan["average_price"])
    direction = plan["profile"]["direction"]
    leverage = float(plan["profile"]["leverage"])
    for label, mode_key, value_key in (("toma parcial", "tp_mode", "tp_value"),
                                       ("activación trailing", "trail_mode", "trail_value")):
        mode = str(setup[mode_key]).upper()
        value = float(setup[value_key])
        price = value if mode == "PRICE" else calc_roi_price(avg, direction, leverage, value)
        if price <= 0:
            raise ValueError(f"El precio calculado para {label} no es válido.")
        if direction == "BUY" and price <= avg:
            raise ValueError(f"La {label} LONG debe quedar por encima del promedio estimado ({fmt(avg,8)}).")
        if direction == "SELL" and price >= avg:
            raise ValueError(f"La {label} SHORT debe quedar por debajo del promedio estimado ({fmt(avg,8)}).")


def process_fill_batch(api: BinanceUSDM, trade: dict[str, Any], qty: float, quote: float,
                       rules: dict[str, Any]) -> tuple[bool, str, float]:
    """Coloca la LIMIT parcial para un bloque de fills.

    El trailing NO se crea por fill: se consolida en una sola orden para todo
    el remanente de la posición, evitando agotar el límite de algos por símbolo.
    """
    if qty <= 0:
        return True, "", 0.0
    slices = trade.setdefault("exit_slices", [])
    sl = trade.get("pending_exit_slice")
    if sl is None:
        avg = quote / qty
        step = rules["step"]
        min_qty = float(rules["min_qty"])
        requested_tp_qty = floor_step(qty * float(trade["setup"]["partial_pct"]) / 100, step)
        tp_qty = requested_tp_qty
        trail_share = floor_step(qty - tp_qty, step)
        note = ""
        # Ajustar por minQty para que ambas patas puedan existir cuando el
        # tamaño lo permite. Si no, proteger el bloque completo con trailing.
        if qty < min_qty:
            return False, f"Fill {fmt(qty,8)} inferior al mínimo {rules['min_qty']}; se acumulará con otros fills.", 0.0
        if tp_qty < min_qty:
            tp_qty = 0.0
            trail_share = floor_step(qty, step)
            note = "El tamaño parcial era inferior al mínimo del mercado; este bloque queda reservado al trailing."
        elif trail_share < min_qty:
            if qty >= 2 * min_qty:
                tp_qty = floor_step(qty - min_qty, step)
                trail_share = floor_step(qty - tp_qty, step)
                note = "Se ajustó la toma parcial para respetar el tamaño mínimo del trailing."
            else:
                tp_qty = 0.0
                trail_share = floor_step(qty, step)
                note = "El bloque es demasiado pequeño para dividirlo; se protege completo con trailing."
        tp_price = None
        if tp_qty >= min_qty:
            tp_price = exit_price_for_fill(trade, avg, "partial")
            if trade["direction"] == "BUY":
                if tp_price <= avg:
                    raise ValueError("La toma parcial LONG debe estar por encima del precio medio del fill.")
                tp_price = quantize_price(tp_price, rules["tick"], "up")
            else:
                if tp_price >= avg:
                    raise ValueError("La toma parcial SHORT debe estar por debajo del precio medio del fill.")
                tp_price = quantize_price(tp_price, rules["tick"], "down")
        n = len(slices) + 1
        tid = trade["id"][-8:].upper()
        sl = {
            "number": n, "fill_qty": qty, "fill_quote": quote, "fill_avg": avg,
            "tp_qty": tp_qty, "tp_price": tp_price,
            "tp_client_id": f"JTP{tid}{n}"[:36], "tp_order_id": None,
            "tp_status": "PENDING" if tp_qty else "OMITIDO",
            "trail_share": trail_share,
            "fallback_note": note, "created_at": int(time.time()),
            "tp_processed_exec_qty": 0.0, "tp_fill_notice_sent": False,
        }
        trade["pending_exit_slice"] = sl
        save_data()  # guardar client ID antes de enviar la orden

    if sl.get("tp_qty", 0) and not sl.get("tp_order_id"):
        try:
            order = api.new_limit(trade["symbol"], "SELL" if trade["direction"] == "BUY" else "BUY",
                                  float(sl["tp_qty"]), float(sl["tp_price"]), sl["tp_client_id"],
                                  trade["position_side"])
        except Exception as send_exc:
            try:
                order = api.order_by_client_id(trade["symbol"], sl["tp_client_id"])
            except Exception:
                sl["tp_last_error"] = str(send_exc)
                save_data()
                return False, f"No se pudo confirmar la orden límite parcial: {send_exc}", 0.0
        sl["tp_order_id"] = int(order["orderId"])
        sl["tp_status"] = order.get("status", "NEW")
        sl["tp_processed_exec_qty"] = float(order.get("executedQty", 0) or 0)
        save_data()

    consumed_qty = float(sl["fill_qty"])
    consumed_quote = float(sl.get("fill_quote", 0))
    slices.append(sl)
    trade["pending_exit_qty"] = max(0.0, float(trade.get("pending_exit_qty", 0)) - consumed_qty)
    trade["pending_exit_quote"] = max(0.0, float(trade.get("pending_exit_quote", 0)) - consumed_quote)
    trade.pop("pending_exit_slice", None)
    save_data()
    note = (f"📌 Gestión parcial registrada para {trade['display_symbol']}\n"
            f"Fill medio: {fmt(sl['fill_avg'],8)} | cantidad: {fmt(sl['fill_qty'],8)}\n"
            f"Orden límite parcial: {fmt(sl['tp_qty'],8)} a {fmt(sl.get('tp_price'),8)}\n"
            f"Remanente planificado para trailing: {fmt(sl['trail_share'],8)}")
    if sl.get("fallback_note"):
        note += "\n⚠️ " + sl["fallback_note"]
    return True, note, consumed_qty


def _tp_open_remaining(api: BinanceUSDM, trade: dict[str, Any]) -> tuple[float, list[str]]:
    """Suma la cantidad aún abierta de todas las LIMIT de toma parcial."""
    remaining = 0.0
    notices: list[str] = []
    terminal = {"FILLED", "CANCELED", "EXPIRED", "REJECTED"}
    for sl in trade.get("exit_slices", []):
        if not sl.get("tp_order_id"):
            continue
        order = api.order(trade["symbol"], int(sl["tp_order_id"]))
        old_status = sl.get("tp_status")
        status = order.get("status", old_status)
        sl["tp_status"] = status
        executed = float(order.get("executedQty", 0) or 0)
        sl["tp_processed_exec_qty"] = executed
        if status not in terminal:
            orig = float(order.get("origQty", sl.get("tp_qty", 0)) or 0)
            remaining += max(0.0, orig - executed)
        if status == "FILLED" and old_status != "FILLED" and not sl.get("tp_fill_notice_sent"):
            notices.append(f"💰 Toma parcial ejecutada en {trade['display_symbol']}: {fmt(executed,8)}")
            sl["tp_fill_notice_sent"] = True
    return remaining, notices


def _trail_activation(api: BinanceUSDM, trade: dict[str, Any], avg_entry: float,
                      rules: dict[str, Any]) -> float | None:
    setup = trade["setup"]
    mode, value = setup["trail_mode"], float(setup["trail_value"])
    activation = value if mode == "PRICE" else calc_roi_price(avg_entry, trade["direction"],
                                                                float(trade["profile"]["leverage"]), value)
    mark = api.mark_price(trade["symbol"])
    state = trade.setdefault("trail_order", {})
    if state.get("activation_reached"):
        return None
    if trade["direction"] == "BUY":
        if activation <= avg_entry:
            raise ValueError("La activación trailing LONG debe estar por encima del precio medio.")
        activation = quantize_price(activation, rules["tick"], "up")
        # En LONG, al tocar/superar activación se conserva ese estado.
        if mark >= activation:
            state["activation_reached"] = True
            return None
        return activation
    if activation >= avg_entry:
        raise ValueError("La activación trailing SHORT debe estar por debajo del precio medio.")
    activation = quantize_price(activation, rules["tick"], "down")
    if mark <= activation:
        state["activation_reached"] = True
        return None
    return activation


def _cancel_remaining_entries(api: BinanceUSDM, trade: dict[str, Any]) -> None:
    for entry in trade.get("entries", []):
        if entry.get("order_id") and entry.get("status") not in {"FILLED", "CANCELED", "EXPIRED", "REJECTED"}:
            try:
                api.cancel_order(trade["symbol"], int(entry["order_id"]))
                entry["status"] = "CANCELED"
            except Exception:
                LOG.exception("No se pudo cancelar la entrada pendiente %s", entry.get("order_id"))


def ensure_aggregated_trailing(api: BinanceUSDM, trade: dict[str, Any]) -> tuple[list[str], bool]:
    """Mantiene un único trailing para todo el remanente no reservado a LIMIT.

    Esto permite hasta 20 entradas sin generar 20 algos condicionales. La pata
    límite es una orden normal por bloque; el trailing es único y se redimensiona
    cuando se confirman nuevos fills.
    """
    notices: list[str] = []
    # Si aún hay fills pendientes o una LIMIT con respuesta ambigua, no crear
    # ni redimensionar trailing: primero debemos conocer exactamente la pata
    # parcial que está viva para evitar solapamientos de reduceOnly.
    if trade.get("pending_exit_slice") or float(trade.get("pending_exit_qty", 0)) > float(trade.get("rules", {}).get("step", "0.00000001")) / 2:
        return notices, False
    pos = api.position(trade["symbol"], trade["direction"])
    position_qty = abs(float(pos.get("positionAmt", 0) or 0))
    rules = trade["rules"]
    min_qty = float(rules["min_qty"])
    state = trade.setdefault("trail_order", {"algo_id": None, "client_id": None, "qty": 0.0,
                                               "activation": None, "status": "NONE", "generation": 0})

    # Si el POST anterior pudo haber sido aceptado pero se perdió su respuesta,
    # recuperar siempre por clientAlgoId. Nunca crear una segunda orden a ciegas.
    if not state.get("algo_id") and state.get("client_id") and state.get("status") == "SUBMISSION_UNKNOWN":
        try:
            recovered = api.algo_order_by_client_id(state["client_id"])
            state["algo_id"] = int(recovered["algoId"])
            state["status"] = recovered.get("algoStatus", "NEW")
            save_data()
        except Exception as exc:
            state["last_error"] = str(exc)
            save_data()
            notices.append(f"⚠️ Binance aún no permite confirmar el trailing anterior. No se enviará otro para evitar duplicarlo; revisa el estado en {trade.get('environment', 'DEMO')}.")
            return notices, False

    tp_remaining, tp_notices = _tp_open_remaining(api, trade)
    notices.extend(tp_notices)
    desired_qty = floor_step(max(0.0, position_qty - tp_remaining), rules["step"])
    if desired_qty < min_qty:
        desired_qty = 0.0

    # Recuperar / inspeccionar el trailing previo.
    if state.get("algo_id"):
        algo = api.algo_order(int(state["algo_id"]))
        status = str(algo.get("algoStatus", state.get("status", "NEW"))).upper()
        state["status"] = status
        if status == "TRIGGERED":
            # El trailing ya se disparó; parar nuevas recompras y no duplicar
            # una salida mientras Binance ejecuta la orden de mercado resultante.
            _cancel_remaining_entries(api, trade)
            state["status"] = status
            save_data()
            return notices, True
        if status == "FINISHED":
            _cancel_remaining_entries(api, trade)
            actual_order_id = algo.get("actualOrderId") or algo.get("actualOrderID")
            if actual_order_id:
                try:
                    child = api.order(trade["symbol"], int(actual_order_id))
                    child_status = str(child.get("status", "")).upper()
                    if child_status in {"NEW", "PARTIALLY_FILLED", "PENDING_NEW"}:
                        state["child_order_id"] = int(actual_order_id)
                        state["child_status"] = child_status
                        save_data()
                        return notices, True
                except Exception:
                    # No crear otra orden mientras no se pueda verificar el hijo.
                    save_data()
                    return notices, True
            # FINISHED significa que el algo ya no está abierto. Se recalcula
            # la posición; si quedó remanente sin límite, puede protegerse de nuevo.
            state["algo_id"] = None
            state["status"] = status
        elif status not in {"CANCELED", "EXPIRED", "REJECTED"}:
            same_qty = abs(float(state.get("qty", 0)) - desired_qty) < float(rules["step"]) / 2
            # Recalcular ROI solo antes de activación; una vez disparado no se modifica.
            avg_entry = float(pos.get("entryPrice", 0) or pos.get("breakEvenPrice", 0) or 0)
            desired_activation = _trail_activation(api, trade, avg_entry, rules) if desired_qty and avg_entry else None
            old_activation = state.get("activation")
            same_activation = ((desired_activation is None and old_activation is None) or
                               (desired_activation is not None and old_activation is not None and
                                abs(float(desired_activation) - float(old_activation)) < float(rules["tick"])))
            if desired_qty <= 0 or (same_qty and same_activation):
                save_data()
                return notices, True
            # Cancelar antes de reemplazar para no dejar dos trailing simultáneos.
            try:
                api.cancel_algo(int(state["algo_id"]))
            except Exception:
                fresh = api.algo_order(int(state["algo_id"]))
                fresh_status = str(fresh.get("algoStatus", "")).upper()
                if fresh_status not in {"CANCELED", "EXPIRED", "REJECTED", "FINISHED"}:
                    raise RuntimeError("No se pudo cancelar el trailing anterior; no se enviará otro para evitar duplicados.")
                if fresh_status == "FINISHED":
                    _cancel_remaining_entries(api, trade)
                    state["status"] = fresh_status
                    save_data()
                    return notices, True
            state["status"] = "CANCELED"
            state["algo_id"] = None
            save_data()
        else:
            state["algo_id"] = None
            state["status"] = status
        
    if desired_qty <= 0:
        save_data()
        return notices, False

    # Refrescar posición y órdenes tras una posible cancelación.
    pos = api.position(trade["symbol"], trade["direction"])
    position_qty = abs(float(pos.get("positionAmt", 0) or 0))
    tp_remaining, more_notices = _tp_open_remaining(api, trade)
    notices.extend(more_notices)
    desired_qty = floor_step(max(0.0, position_qty - tp_remaining), rules["step"])
    if desired_qty < min_qty:
        save_data()
        return notices, False
    avg_entry = float(pos.get("entryPrice", 0) or pos.get("breakEvenPrice", 0) or 0)
    if avg_entry <= 0:
        raise RuntimeError("Binance no devolvió el precio medio de la posición; trailing no enviado.")
    activation = _trail_activation(api, trade, avg_entry, rules)
    if activation is None:
        pending_before = any(e.get("status") not in {"FILLED", "CANCELED", "EXPIRED", "REJECTED", "CANCELED_BEFORE_SUBMIT"}
                             for e in trade.get("entries", []))
        _cancel_remaining_entries(api, trade)
        if pending_before:
            notices.append("🎯 Se alcanzó la activación del trailing. Se cancelaron las recompras pendientes para no aumentar la exposición mientras protege ganancias.")
    state["generation"] = int(state.get("generation", 0)) + 1
    client_id = f"JHT{trade['id'][-8:].upper()}{state['generation']}"[:36]
    state.update({"client_id": client_id, "qty": desired_qty, "activation": activation,
                  "status": "SUBMISSION_UNKNOWN", "algo_id": None})
    save_data()  # idempotencia antes del POST
    side_exit = "SELL" if trade["direction"] == "BUY" else "BUY"
    try:
        response = api.new_trailing(trade["symbol"], side_exit, desired_qty, activation,
                                    float(trade["setup"]["callback_rate"]), client_id,
                                    trade["position_side"])
    except Exception as send_exc:
        try:
            response = api.algo_order_by_client_id(client_id)
        except Exception:
            state["last_error"] = str(send_exc)
            save_data()
            raise RuntimeError(f"No se pudo confirmar el trailing agregado: {send_exc}") from send_exc
    state["algo_id"] = int(response["algoId"])
    state["status"] = response.get("algoStatus", "NEW")
    state.pop("last_error", None)
    save_data()
    notices.append(
        f"🛡️ Trailing agregado en {trade['display_symbol']}\n"
        f"Cantidad protegida: {fmt(desired_qty,8)}\n"
        f"Activación: {fmt(activation,8) if activation is not None else 'inmediata desde el precio vigente'}\n"
        f"Retroceso: {trade['setup']['callback_rate']}%"
    )
    return notices, True


def cancel_other_exit_orders(api: BinanceUSDM, trade: dict[str, Any]) -> list[str]:
    """Cancela órdenes de salida hermanas después de dispararse el stop fijo."""
    notes: list[str] = []
    _cancel_remaining_entries(api, trade)
    for sl in trade.get("exit_slices", []):
        oid = sl.get("tp_order_id")
        if not oid:
            continue
        try:
            current = api.order(trade["symbol"], int(oid))
            if current.get("status") not in {"FILLED", "CANCELED", "EXPIRED", "REJECTED"}:
                api.cancel_order(trade["symbol"], int(oid))
            sl["tp_status"] = "CANCELED"
        except Exception as exc:
            notes.append(f"⚠️ No se pudo cancelar la toma parcial {oid}: {exc}")
    trail = trade.get("trail_order", {})
    aid = trail.get("algo_id")
    if aid and str(trail.get("status", "")).upper() not in {"FINISHED", "TRIGGERED", "CANCELED", "EXPIRED", "REJECTED"}:
        try:
            api.cancel_algo(int(aid))
            trail["status"] = "CANCELED"
        except Exception as exc:
            notes.append(f"⚠️ No se pudo cancelar trailing {aid}: {exc}")
    save_data()
    return notes


def ensure_fixed_stop_loss(api: BinanceUSDM, trade: dict[str, Any]) -> tuple[list[str], bool]:
    """Crea/actualiza el STOP_MARKET closePosition según la fórmula del Excel.

    Devuelve (avisos, stop_disparado). El cierre se asocia exclusivamente a
    positionSide LONG o SHORT, no toca la pata opuesta en Hedge Mode.
    """
    notes: list[str] = []
    state = trade.setdefault("stop_loss_order", {
        "algo_id": None, "client_id": None, "trigger_price": None,
        "status": "NONE", "generation": 0,
    })
    position = api.position(trade["symbol"], trade["direction"])
    position_qty = abs(float(position.get("positionAmt", 0) or 0))
    if position_qty <= 0:
        # No hay posición de esta pata. Si quedó un stop antiguo activo, cancelarlo.
        if state.get("algo_id") and str(state.get("status", "")).upper() not in {"FINISHED", "TRIGGERED", "CANCELED", "EXPIRED", "REJECTED"}:
            try:
                api.cancel_algo(int(state["algo_id"]))
            except Exception:
                pass
        state["status"] = "NO_POSITION"
        save_data()
        return notes, False

    # Recupera una orden MARKET de emergencia si la respuesta se perdió.
    if not state.get("market_order_id") and state.get("client_id") and state.get("status") == "MARKET_CLOSE_UNKNOWN":
        try:
            recovered_market = api.order_by_client_id(trade["symbol"], state["client_id"])
            state["market_order_id"] = int(recovered_market["orderId"])
            state["status"] = recovered_market.get("status", "NEW")
            save_data()
        except Exception as exc:
            state["last_error"] = str(exc)
            save_data()
            notes.append("🚨 Cierre MARKET de emergencia no confirmado. No se repetirá a ciegas; verifica la pata en Binance.")
            return notes, True

    if state.get("market_order_id"):
        try:
            market_order = api.order(trade["symbol"], int(state["market_order_id"]))
            market_status = str(market_order.get("status", "NEW")).upper()
            state["status"] = market_status
            if market_status in {"NEW", "PARTIALLY_FILLED", "PENDING_NEW"}:
                save_data()
                return notes, True
            latest_pos = api.position(trade["symbol"], trade["direction"])
            latest_qty = abs(float(latest_pos.get("positionAmt", 0) or 0))
            if latest_qty <= 0:
                save_data()
                return notes, True
            # La MARKET anterior terminó pero quedó cantidad residual: se
            # permite un nuevo cierre por el remanente con client ID distinto.
            state.pop("market_order_id", None)
            state["status"] = "MARKET_CLOSE_RETRY"
            position = latest_pos
            position_qty = latest_qty
            save_data()
        except Exception as exc:
            state["last_error"] = str(exc)
            save_data()
            notes.append(f"🚨 No se pudo verificar el cierre MARKET previo: {exc}")
            return notes, True

    # Recuperación idempotente si el POST pudo llegar a Binance y se perdió la respuesta.
    if not state.get("algo_id") and state.get("client_id") and state.get("status") == "SUBMISSION_UNKNOWN":
        try:
            recovered = api.algo_order_by_client_id(state["client_id"])
            state["algo_id"] = int(recovered["algoId"])
            state["status"] = recovered.get("algoStatus", "NEW")
            save_data()
        except Exception as exc:
            state["last_error"] = str(exc)
            save_data()
            notes.append("🚨 No se puede confirmar el stop anterior. No se enviará uno duplicado; revisa Binance.")
            return notes, True

    if state.get("algo_id"):
        try:
            current = api.algo_order(int(state["algo_id"]))
            status = str(current.get("algoStatus", state.get("status", "NEW"))).upper()
            state["status"] = status
            if status in {"TRIGGERED", "FINISHED"}:
                trade["exit_reason"] = "STOP_LOSS"
                notes.extend(cancel_other_exit_orders(api, trade))
                save_data()
                return notes, True
            if status not in {"CANCELED", "EXPIRED", "REJECTED"}:
                pass
            else:
                state["algo_id"] = None
        except Exception as exc:
            notes.append(f"⚠️ No se pudo consultar el stop existente: {exc}")
            return notes, True

    avg_entry = float(position.get("entryPrice", 0) or position.get("breakEvenPrice", 0) or 0)
    if avg_entry <= 0:
        notes.append("🚨 Binance no devolvió el precio promedio para actualizar el stop.")
        return notes, True
    prof = trade["profile"]
    stop_roi_pct = float(prof["stop_roi_pct"])
    leverage = float(prof["leverage"])
    raw_trigger = calc_excel_stop_price(avg_entry, trade["direction"], leverage, stop_roi_pct)
    tick = trade["rules"]["tick"]
    trigger = quantize_price(raw_trigger, tick, "up" if trade["direction"] == "BUY" else "down")
    mark = api.mark_price(trade["symbol"])

    # El movimiento adverso es stop_roi_pct / leverage / 100.
    # Por ejemplo, BTC/ETH 300% a 100x = 3%; altcoin 80% a 20x = 4%.
    # La liquidación depende del margen disponible y del maintenance tier; si
    # Binance reporta una liquidación antes del stop, avisar claramente.
    liq_price = float(position.get("liquidationPrice", 0) or 0)
    stop_behind_liquidation = bool(
        liq_price > 0 and (
            (trade["direction"] == "BUY" and trigger <= liq_price) or
            (trade["direction"] == "SELL" and trigger >= liq_price)
        )
    )
    if stop_behind_liquidation and not state.get("liquidation_warning_sent"):
        notes.append(
            f"🚨 ALERTA DE LIQUIDACIÓN {trade['display_symbol']} ({trade['position_side']})\n"
            f"Precio de liquidación informado por Binance: {fmt(liq_price,8)}\n"
            f"Stop del Excel: {fmt(trigger,8)}. El stop queda más allá del precio de liquidación; Binance podría liquidar antes de que se ejecute el STOP_MARKET. Revisa margen, tamaño y precio de liquidación inmediatamente."
        )
        state["liquidation_warning_sent"] = True
    elif not stop_behind_liquidation:
        state["liquidation_warning_sent"] = False

    # Si el precio ya cruzó el stop antes de poder crear la protección,
    # liquidar únicamente esta pata Hedge con una MARKET idempotente.
    crossed = (mark <= trigger) if trade["direction"] == "BUY" else (mark >= trigger)
    if crossed:
        _cancel_remaining_entries(api, trade)
        notes.extend(cancel_other_exit_orders(api, trade))
        generation = int(state.get("generation", 0)) + 1
        cid = f"JSLM{trade['id'][-8:].upper()}{generation}"[:36]
        state.update({"client_id": cid, "status": "MARKET_CLOSE_UNKNOWN", "generation": generation,
                      "trigger_price": trigger})
        save_data()
        try:
            response = api.new_market_close(
                trade["symbol"], "SELL" if trade["direction"] == "BUY" else "BUY",
                floor_step(position_qty, trade["rules"]["step"]), trade["position_side"], cid)
        except Exception as send_exc:
            try:
                response = api.order_by_client_id(trade["symbol"], cid)
            except Exception:
                state["last_error"] = str(send_exc)
                save_data()
                notes.append(f"🚨 El precio ya atravesó el stop ({fmt(trigger,8)}), pero no se pudo confirmar el cierre MARKET: {send_exc}")
                return notes, True
        state["market_order_id"] = int(response["orderId"])
        state["status"] = response.get("status", "NEW")
        trade["exit_reason"] = "STOP_LOSS_MARKET"
        save_data()
        notes.append(f"🛑 El precio ya había atravesado el stop {fmt(trigger,8)}. Se envió MARKET para cerrar {trade['position_side']}.")
        return notes, True

    old_trigger = state.get("trigger_price")
    if state.get("algo_id") and state.get("status") not in {"CANCELED", "EXPIRED", "REJECTED", "NO_POSITION"}:
        same = old_trigger is not None and abs(float(old_trigger) - trigger) < float(tick) / 2
        if same:
            return notes, False
        # El promedio cambió por un nuevo fill. Reemplazar el stop close-all.
        try:
            api.cancel_algo(int(state["algo_id"]))
        except Exception as exc:
            fresh = api.algo_order(int(state["algo_id"]))
            fresh_status = str(fresh.get("algoStatus", "")).upper()
            if fresh_status in {"TRIGGERED", "FINISHED"}:
                trade["exit_reason"] = "STOP_LOSS"
                notes.extend(cancel_other_exit_orders(api, trade))
                return notes, True
            if fresh_status not in {"CANCELED", "EXPIRED", "REJECTED"}:
                raise RuntimeError(f"No se pudo reemplazar el stop anterior: {exc}")
        state["status"] = "CANCELED"
        state["algo_id"] = None
        save_data()

    generation = int(state.get("generation", 0)) + 1
    cid = f"JSL{trade['id'][-8:].upper()}{generation}"[:36]
    side_exit = "SELL" if trade["direction"] == "BUY" else "BUY"
    state.update({"client_id": cid, "trigger_price": trigger, "generation": generation,
                  "status": "SUBMISSION_UNKNOWN", "algo_id": None,
                  "stop_roi_pct": stop_roi_pct, "leverage": leverage, "avg_entry": avg_entry})
    save_data()
    try:
        response = api.new_stop_loss(trade["symbol"], side_exit, trigger, cid, trade["position_side"])
    except Exception as send_exc:
        try:
            response = api.algo_order_by_client_id(cid)
        except Exception:
            state["last_error"] = str(send_exc)
            save_data()
            raise RuntimeError(f"No se pudo confirmar el STOP_MARKET: {send_exc}") from send_exc
    state["algo_id"] = int(response["algoId"])
    state["status"] = response.get("algoStatus", "NEW")
    state.pop("last_error", None)
    save_data()
    notes.append(
        f"🛑 Stop Loss {trade['display_symbol']} ({trade['position_side']})\n"
        f"Promedio real: {fmt(avg_entry,8)} | Riesgo Excel: {stop_roi_pct:g}% ROI / {leverage:g}x\n"
        f"Movimiento adverso: {stop_roi_pct/leverage:.4f}% | Activación: {fmt(trigger,8)}"
    )
    return notes, False

# -------------------------------- Telegram -----------------------------------
def allowed(update: Update) -> bool:
    user = update.effective_user
    if not user:
        return False
    raw = os.getenv("TELEGRAM_ALLOWED_USER_IDS", "").strip()
    if raw:
        try:
            return user.id in {int(x.strip()) for x in raw.split(",") if x.strip()}
        except ValueError:
            return False
    owner = DATA.get("owner_user_id")
    if owner is None:
        DATA["owner_user_id"] = user.id
        save_data()
        return True
    return int(owner) == user.id

def menu_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📨 Reenviar / analizar señal", callback_data="await_signal")],
        [InlineKeyboardButton("📈 Estado de operaciones", callback_data="status")],
        [InlineKeyboardButton("⚙️ Configuración", callback_data="config")],
        [InlineKeyboardButton(f"🔌 Binance {configured_binance_mode()}", callback_data="binance")],
    ])

def signal_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🧮 Calcular con configuración guardada", callback_data="calc")],
        [InlineKeyboardButton("🎯 Configurar salida para esta señal", callback_data="configure_exits")],
        [InlineKeyboardButton("❌ Descartar señal", callback_data="discard")],
    ])

def report_kb() -> InlineKeyboardMarkup:
    mode = configured_binance_mode()
    label = "⚠️ Preparar envío de ÓRDENES REALES" if mode == "LIVE" else "🧪 Enviar entradas a Binance DEMO"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(label, callback_data="execute")],
        [InlineKeyboardButton("🏠 Menú", callback_data="menu")],
    ])

def live_confirmation_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⚠️ CONFIRMO: ENVIAR ÓRDENES REALES", callback_data="confirm_live_execute")],
        [InlineKeyboardButton("Cancelar", callback_data="cancel_live_execute")],
    ])

SETUP_STEPS = ["capital", "btc_count", "btc_operate", "btc_stop_roi_pct", "btc_alloc",
               "alt_count", "alt_operate", "alt_leverage", "alt_stop_roi_pct", "alt_alloc",
               "partial_pct", "tp_mode", "tp_value", "trail_mode", "trail_value", "callback_rate"]

def wizard_prompt(step: str) -> str:
    prompts = {
        "capital": "Configuración 1/16 — Capital total disponible en USDT. Ejemplo: 1000",
        "btc_count": "2/16 — Entradas BTC/ETH (1 a 20). Recomendación Excel: 3, 4, 5 o 6.",
        "btc_operate": "3/16 — % del capital destinado por operación BTC/ETH. Escribe DEFAULT para usar el valor de Excel según la plantilla.",
        "btc_stop_roi_pct": "4/16 — Stop Loss BTC/ETH en ROI negativo. Escribe el valor positivo que representa el riesgo (300 = −300% ROI). DEFAULT = 300.",
        "btc_alloc": "5/16 — Porcentajes de cada entrada BTC/ETH separados por coma y sumando 100, o DEFAULT para distribución Excel.",
        "alt_count": "6/16 — Entradas para ALTCOINS (1 a 20). Recomendación Excel: 3, 4 o 6.",
        "alt_operate": "7/16 — % de capital por operación ALTCOINS. DEFAULT conserva los valores por plantilla Excel.",
        "alt_leverage": "8/16 — Apalancamiento ALTCOINS: elige un entero entre 20x y 30x. DEFAULT = 20x.",
        "alt_stop_roi_pct": "9/16 — Stop Loss ALTCOINS en ROI negativo: elige entre 80% y 120%. Ejemplo: 100 = −100% ROI. DEFAULT = 120.",
        "alt_alloc": "10/16 — Porcentajes de entradas ALTCOINS, separados por coma y sumando 100, o DEFAULT.",
        "partial_pct": "11/16 — ¿Qué porcentaje del tamaño ejecutado tomará la orden límite parcial? Ejemplo: 50",
        "tp_mode": "12/16 — ¿La toma parcial se define por PRECIO o por ROI? Escribe PRECIO o ROI.",
        "tp_value": "13/16 — Valor de la toma parcial: precio exacto USDT o ROI % según elegiste.",
        "trail_mode": "14/16 — ¿La activación del trailing se define por PRECIO o por ROI? Escribe PRECIO o ROI.",
        "trail_value": "15/16 — Valor de activación del trailing: precio exacto o ROI % según elegiste.",
        "callback_rate": "16/16 — Retroceso del trailing en % (Binance USDⓈ-M admite 0.1 a 20). Ejemplo: 1",
    }
    return prompts[step]

def initial_setup_data() -> dict[str, Any]:
    return {"kind": "setup", "steps": SETUP_STEPS[:], "step": 0, "values": {}}

async def start_exit_wizard(chat_id: int, uid: str, context: ContextTypes.DEFAULT_TYPE) -> None:
    if uid not in DATA.get("pending", {}):
        await context.bot.send_message(chat_id, "No hay señal pendiente para configurar.")
        return
    steps = ["tp_mode", "tp_value", "trail_mode", "trail_value"]
    DATA["wizard"][uid] = {"kind": "exit", "steps": steps, "step": 0, "values": {}}
    save_data()
    prompts = {
        "tp_mode": "SALIDA DE ESTA SEÑAL 1/4 — Toma parcial por PRECIO o ROI? Escribe PRECIO o ROI.",
        "tp_value": "2/4 — Escribe el precio exacto USDT o el ROI % de la orden límite parcial.",
        "trail_mode": "3/4 — Activación del trailing por PRECIO o ROI? Escribe PRECIO o ROI.",
        "trail_value": "4/4 — Escribe el precio de activación o ROI % del trailing.",
    }
    await context.bot.send_message(chat_id, prompts[steps[0]], reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Cancelar", callback_data="cancel_wizard")]]))

async def start_wizard(chat_id: int, uid: str, context: ContextTypes.DEFAULT_TYPE) -> None:
    DATA["wizard"][uid] = initial_setup_data()
    save_data()
    await context.bot.send_message(chat_id, wizard_prompt(SETUP_STEPS[0]),
                                   reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Cancelar", callback_data="cancel_wizard")]]))

async def consume_wizard(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    uid = str(update.effective_user.id)
    wizard = DATA.get("wizard", {}).get(uid)
    if not wizard:
        return False
    msg = update.effective_message
    text = (msg.text or msg.caption or "").strip()
    steps = wizard.get("steps", SETUP_STEPS)
    step = steps[int(wizard["step"])]
    vals = wizard["values"]
    try:
        if step == "capital":
            v = parse_num(text)
            if v <= 0: raise ValueError("El capital debe ser mayor que cero.")
            vals[step] = v
        elif step.endswith("count"):
            v = int(text)
            if not 1 <= v <= MAX_ENTRIES: raise ValueError("Usa un número entre 1 y 20.")
            vals[step] = v
        elif step.endswith("operate"):
            if text.upper() == "DEFAULT": vals[step] = None
            else:
                v = parse_num(text)
                if not 0 < v <= 100:
                    raise ValueError("El porcentaje debe ser mayor que 0 y no superar 100.")
                vals[step] = v
        elif step == "alt_leverage":
            v = 20.0 if text.upper() == "DEFAULT" else parse_num(text)
            if not v.is_integer() or not 20 <= v <= 30:
                raise ValueError("Para altcoins el apalancamiento debe ser un entero entre 20x y 30x.")
            vals[step] = int(v)
        elif step == "btc_stop_roi_pct":
            v = 300.0 if text.upper() == "DEFAULT" else parse_num(text)
            if not 1 <= v <= 1000:
                raise ValueError("El riesgo ROI BTC/ETH debe estar entre 1% y 1000%. Ejemplo: 300.")
            vals[step] = v
        elif step == "alt_stop_roi_pct":
            v = 120.0 if text.upper() == "DEFAULT" else parse_num(text)
            if not 80 <= v <= 120:
                raise ValueError("El Stop Loss de altcoins debe estar entre 80% y 120% ROI.")
            vals[step] = v
        elif step.endswith("alloc"):
            count = vals["btc_count"] if step == "btc_alloc" else vals["alt_count"]
            kind = "btc_eth" if step == "btc_alloc" else "altcoin"
            vals[step] = None if text.upper() in {"DEFAULT", "PREDETERMINADO", "EXCEL"} else pct_allocations(text, count, default_weights(kind, count, "BUY"))
        elif step == "partial_pct":
            v = parse_num(text)
            if not 1 <= v <= 99: raise ValueError("Usa un porcentaje entre 1 y 99.")
            vals[step] = v
        elif step in {"tp_mode", "trail_mode"}:
            v = text.upper()
            if v not in {"ROI", "PRECIO"}: raise ValueError("Escribe ROI o PRECIO.")
            vals[step] = v
        elif step in {"tp_value", "trail_value"}:
            v = parse_num(text)
            if v <= 0: raise ValueError("El valor debe ser mayor que cero.")
            mode_key = "tp_mode" if step == "tp_value" else "trail_mode"
            if vals[mode_key] == "ROI" and v > 10000: raise ValueError("ROI demasiado alto.")
            vals[step] = v
        elif step == "callback_rate":
            v = parse_num(text)
            if not 0.1 <= v <= 20: raise ValueError("Binance USDⓈ-M Futures admite callback entre 0.1 y 20 % para trailing.")
            vals[step] = v
    except (ValueError, KeyError) as exc:
        await msg.reply_text(f"Dato no válido: {exc}\n\n{wizard_prompt(step)}")
        return True
    wizard["step"] += 1
    if wizard["step"] < len(steps):
        save_data()
        next_step = steps[wizard["step"]]
        if wizard.get("kind") == "exit":
            exit_prompts = {
                "tp_mode": "1/4 — Toma parcial por PRECIO o ROI? Escribe PRECIO o ROI.",
                "tp_value": "2/4 — Precio exacto USDT o ROI % de la orden límite parcial.",
                "trail_mode": "3/4 — Activación del trailing por PRECIO o ROI? Escribe PRECIO o ROI.",
                "trail_value": "4/4 — Precio de activación o ROI % del trailing.",
            }
            await msg.reply_text(exit_prompts[next_step])
        else:
            await msg.reply_text(wizard_prompt(next_step))
        return True

    if wizard.get("kind") == "exit":
        pending = DATA.get("pending", {}).get(uid)
        if not pending:
            DATA["wizard"].pop(uid, None); save_data()
            await msg.reply_text("La señal ya no está pendiente. Vuelve a reenviarla.", reply_markup=menu_kb())
            return True
        pending["exit_settings"] = {
            "tp_mode": vals["tp_mode"], "tp_value": vals["tp_value"],
            "trail_mode": vals["trail_mode"], "trail_value": vals["trail_value"],
        }
        DATA["wizard"].pop(uid, None)
        save_data()
        effective_setup = {**DATA["setup"], **pending["exit_settings"]}
        await msg.reply_text("✅ Salidas personalizadas para esta señal. El porcentaje parcial y callback siguen usando tu configuración guardada.\n\n" + format_plan(pending["signal"], pending["plan"], effective_setup), reply_markup=report_kb())
        return True

    DATA["setup"] = {
        "complete": True, "capital": vals["capital"],
        "btc_count": vals["btc_count"], "btc_operate": vals["btc_operate"],
        "btc_leverage": 100, "btc_stop_roi_pct": vals["btc_stop_roi_pct"],
        "btc_allocations": vals["btc_alloc"],
        "alt_count": vals["alt_count"], "alt_operate": vals["alt_operate"],
        "alt_leverage": vals["alt_leverage"], "alt_stop_roi_pct": vals["alt_stop_roi_pct"],
        "alt_allocations": vals["alt_alloc"],
        "partial_pct": vals["partial_pct"], "tp_mode": vals["tp_mode"], "tp_value": vals["tp_value"],
        "trail_mode": vals["trail_mode"], "trail_value": vals["trail_value"],
        "callback_rate": vals["callback_rate"],
    }
    DATA["wizard"].pop(uid, None)
    save_data()
    await msg.reply_text("✅ Configuración guardada. Se conservará después de reiniciar.\n\n" + setup_summary(DATA["setup"]), reply_markup=menu_kb())
    return True

def setup_summary(s: dict[str, Any]) -> str:
    btc_stop = float(s.get("btc_stop_roi_pct", 300.0))
    alt_stop = float(s.get("alt_stop_roi_pct", 120.0))
    alt_lev = int(s.get("alt_leverage") or 20)
    return (f"Capital: {s['capital']} USDT\nBTC/ETH: {s['btc_count']} entradas | % operar: {s['btc_operate'] if s['btc_operate'] is not None else 'Excel'} | leverage: 100x | Stop: -{btc_stop:g}% ROI = {btc_stop/100:.4g}% de movimiento del precio | distribución: {'Excel por dirección' if not s.get('btc_allocations') else 'personalizada'}\n"
            f"Altcoins: {s['alt_count']} entradas | % operar: {s['alt_operate'] if s['alt_operate'] is not None else 'Excel'} | leverage: {alt_lev}x | Stop: -{alt_stop:g}% ROI = {alt_stop/alt_lev:.4g}% de movimiento del precio | distribución: {'Excel por dirección' if not s.get('alt_allocations') else 'personalizada'}\n"
            f"Toma parcial: {s['partial_pct']}% por {s['tp_mode']} {s['tp_value']}\nTrailing: activación {s['trail_mode']} {s['trail_value']} | retroceso {s['callback_rate']}%")

async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        await update.effective_message.reply_text("No autorizado.")
        return
    if DATA.get("setup", {}).get("complete"):
        await update.effective_message.reply_text("Bot listo. Tu configuración permanece guardada.", reply_markup=menu_kb())
    else:
        await start_wizard(update.effective_chat.id, str(update.effective_user.id), context)

async def config_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        await update.effective_message.reply_text("No autorizado.")
        return
    await start_wizard(update.effective_chat.id, str(update.effective_user.id), context)

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        await update.effective_message.reply_text("No autorizado.")
        return
    await update.effective_message.reply_text(
        "Reenvía un mensaje que contenga un par USDT, BUY/SELL y Rango.\n"
        "Solo se extraen símbolo, dirección y rango. TP, SL y leverage de la señal se ignoran.\n"
        f"Entorno configurado: {mode_label()}. En LIVE se exige confirmación adicional en Telegram. La señal aporta solo símbolo, dirección y rango; ignora sus TP/SL/leverage. El bot trabaja en Hedge Mode, configura STOP_MARKET según perfil y gestiona toma parcial LIMIT + trailing.", reply_markup=menu_kb())

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not msg or not update.effective_user:
        return
    if not allowed(update):
        await msg.reply_text("No autorizado.")
        return
    if await consume_wizard(update, context):
        return
    if not DATA.get("setup", {}).get("complete"):
        await msg.reply_text("Primero configura el bot con /start.")
        return
    text = msg.text or msg.caption or ""
    try:
        signal = parse_signal(text)
        plan = build_plan(signal, DATA["setup"])
    except Exception as exc:
        await msg.reply_text(f"No pude leer la señal: {exc}\nSolo necesito par, BUY/SELL y rango.")
        return
    uid = str(update.effective_user.id)
    DATA["pending"][uid] = {"signal": signal, "plan": plan}
    DATA.get("live_confirm", {}).pop(uid, None)
    # JSON-safe: plan stored as a dict of plain numbers/strings.
    save_data()
    await msg.reply_text(format_plan(signal, plan, DATA["setup"]), reply_markup=signal_kb())

def pending_signal_digest(pending: dict[str, Any]) -> str:
    packed = json.dumps(pending, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(packed.encode("utf-8")).hexdigest()

async def execute_pending_signal(chat_id: int, uid: str, context: ContextTypes.DEFAULT_TYPE) -> None:
    p = DATA.get("pending", {}).get(uid)
    if not p:
        await context.bot.send_message(chat_id, "No hay señal pendiente para ejecutar.", reply_markup=menu_kb())
        return
    mode = configured_binance_mode()
    if mode == "LIVE" and not live_trading_allowed():
        await context.bot.send_message(chat_id, "⛔ El modo LIVE está bloqueado. Configura ALLOW_LIVE_TRADING=YES en .env y reinicia el bot.", reply_markup=menu_kb())
        return
    await context.bot.send_message(chat_id, f"Validando y enviando entradas a Binance USDⓈ-M Futures ({mode_label(mode)})…")
    try:
        trade_setup = {**DATA["setup"], **p.get("exit_settings", {})}
        validate_exit_settings(p["plan"], trade_setup)
        result = await asyncio.to_thread(submit_trade, p["signal"], p["plan"], trade_setup, chat_id)
        DATA["trades"][result["id"]] = result
        DATA["pending"].pop(uid, None)
        DATA.get("live_confirm", {}).pop(uid, None)
        DATA["last_signal"][uid] = p["signal"]
        save_data()
        summary = execution_summary(result)
        if result.get("submission_warning"):
            summary += "\n\n⚠️ " + result["submission_warning"]
        await context.bot.send_message(chat_id, summary, reply_markup=menu_kb())
    except Exception as exc:
        LOG.exception("No se pudo enviar el plan a Binance")
        await context.bot.send_message(chat_id, f"⛔ No se pudo completar el envío de órdenes.\n{exc}\n\nRevisa el estado de las órdenes en Binance antes de volver a intentarlo; no repitas la señal a ciegas.", reply_markup=menu_kb())

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q:
        return
    if not allowed(update):
        await q.answer("No autorizado", show_alert=True)
        return
    await q.answer()
    uid = str(q.from_user.id)
    action = q.data
    if action == "await_signal":
        await q.edit_message_text("📨 Reenvíame la señal o pega su texto. Solo necesito moneda, BUY/SELL y rango.", reply_markup=menu_kb())
    elif action == "config":
        await start_wizard(q.message.chat_id, uid, context)
        await q.edit_message_text("Iniciando configuración. Responde las preguntas en este chat.")
    elif action == "cancel_wizard":
        DATA["wizard"].pop(uid, None); save_data()
        await q.edit_message_text("Configuración cancelada.", reply_markup=menu_kb())
    elif action == "discard":
        DATA["pending"].pop(uid, None); save_data()
        await q.edit_message_text("Señal descartada.", reply_markup=menu_kb())
    elif action == "calc":
        p = DATA.get("pending", {}).get(uid)
        if not p:
            await q.edit_message_text("No hay señal pendiente.", reply_markup=menu_kb()); return
        effective_setup = {**DATA["setup"], **p.get("exit_settings", {})}
        await q.edit_message_text(format_plan(p["signal"], p["plan"], effective_setup), reply_markup=report_kb())
    elif action == "configure_exits":
        await start_exit_wizard(q.message.chat_id, uid, context)
        await q.edit_message_text("🎯 Configurar toma parcial y trailing para esta señal. Responde las cuatro preguntas que te envié.")
    elif action == "execute":
        p = DATA.get("pending", {}).get(uid)
        if not p:
            await q.edit_message_text("No hay señal pendiente.", reply_markup=menu_kb()); return
        mode = configured_binance_mode()
        if mode == "LIVE":
            if not live_trading_allowed():
                await q.edit_message_text("⛔ Modo LIVE bloqueado. En .env configura ALLOW_LIVE_TRADING=YES y reinicia el bot.", reply_markup=menu_kb()); return
            DATA.setdefault("live_confirm", {})[uid] = pending_signal_digest(p)
            save_data()
            await q.edit_message_text(
                "⚠️ CONFIRMACIÓN DE OPERACIÓN REAL\n\n"
                f"{p['signal']['display_symbol']} | {'LONG' if p['signal']['direction']=='BUY' else 'SHORT'}\n"
                f"Entradas límite: {len(p['plan']['entries'])}\n"
                f"Margen estimado: {p['plan']['total_margin']:.4f} USDT\n\n"
                "Al confirmar, Binance recibirá órdenes reales con fondos reales. Verifica símbolo, precios, apalancamiento y margen. Esta acción no es reversible una vez aceptadas las órdenes.",
                reply_markup=live_confirmation_kb())
        else:
            await q.edit_message_text("Enviando entradas a Binance Futures Demo…")
            await execute_pending_signal(q.message.chat_id, uid, context)
    elif action == "confirm_live_execute":
        p = DATA.get("pending", {}).get(uid)
        ticket = DATA.get("live_confirm", {}).get(uid)
        if not p or not ticket or ticket != pending_signal_digest(p):
            DATA.get("live_confirm", {}).pop(uid, None); save_data()
            await q.edit_message_text("La señal cambió o venció la confirmación. Analízala de nuevo.", reply_markup=menu_kb()); return
        if configured_binance_mode() != "LIVE" or not live_trading_allowed():
            await q.edit_message_text("El modo LIVE no está habilitado en .env.", reply_markup=menu_kb()); return
        DATA.get("live_confirm", {}).pop(uid, None); save_data()
        await q.edit_message_text("Confirmación recibida. Enviando órdenes reales a Binance…")
        await execute_pending_signal(q.message.chat_id, uid, context)
    elif action == "cancel_live_execute":
        DATA.get("live_confirm", {}).pop(uid, None); save_data()
        await q.edit_message_text("Envío real cancelado. No se enviaron órdenes.", reply_markup=menu_kb())
    elif action == "status":
        await q.edit_message_text(status_text(), reply_markup=menu_kb())
    elif action == "binance":
        await q.edit_message_text(f"Comprobando cuenta Binance Futures {configured_binance_mode()}…")
        try:
            report = await asyncio.to_thread(binance_status)
            await context.bot.send_message(q.message.chat_id, report, reply_markup=menu_kb())
        except Exception as exc:
            await context.bot.send_message(q.message.chat_id, f"No se pudo verificar Binance {configured_binance_mode()}: {exc}", reply_markup=menu_kb())
    elif action == "menu":
        await q.edit_message_text("Menú principal", reply_markup=menu_kb())

# -------------------------- Construcción de plan ----------------------------
def format_plan(signal: dict[str, Any], plan: dict[str, Any], setup: dict[str, Any]) -> str:
    p = plan["profile"]
    lines = [f"📨 SEÑAL: {signal['display_symbol']} | {'LONG' if signal['direction']=='BUY' else 'SHORT'}",
             f"Perfil Excel: {p['sheet']} | margen: {p['margin_type']} | leverage: {p['leverage']:g}x",
             f"Rango detectado: {fmt(signal['min_price'], p['price_decimals'])} – {fmt(signal['max_price'], p['price_decimals'])}",
             f"Rango aplicado a entradas: {plan['range_fraction']*100:.4f}%",
             f"Capital: {setup['capital']} USDT | % operar: {p['operate_pct']:g}% | margen aprox.: {plan['total_margin']:.4f} USDT",
             f"Entorno de envío: {mode_label()}",
             f"Distribución: {', '.join(f'{x:g}%' for x in p['allocations'])}", "", "ENTRADAS LIMIT PLANIFICADAS"]
    for e in plan["entries"]:
        lines.append(f"{e['number']:02d}. {fmt(e['price'], p['price_decimals'])} | {e['weight']:g}% | margen {e['margin']:.6f} | qty Excel {fmt(e['quantity_excel'], p['quantity_decimals'])}")
    estimated_stop = calc_excel_stop_price(plan["average_price"], signal["direction"],
                                           p["leverage"], p["stop_roi_pct"])
    adverse_pct = p["stop_roi_pct"] / p["leverage"]
    lines += ["", f"Promedio ponderado estimado: {fmt(plan['average_price'], p['price_decimals'])}",
              f"Stop Loss CONFIGURABLE: -{p['stop_roi_pct']:g}% ROI / {p['leverage']:g}x = {adverse_pct:.4f}% adverso | precio aprox. {fmt(estimated_stop, p['price_decimals'])}",
              ("⚠️ El Stop Loss se calcula desde el promedio real de la pata. La liquidación depende del margen y del tramo de mantenimiento; el bot avisará si Binance informa que la liquidación está antes del stop."),
              f"Toma parcial: {setup['partial_pct']}% por {setup['tp_mode']} {setup['tp_value']}",
              f"Trailing: activar por {setup['trail_mode']} {setup['trail_value']} | callback {setup['callback_rate']}%",
              "", "Los TP/SL/leverage del mensaje de Telegram se ignoran.",
              f"La toma parcial LIMIT se coloca tras cada fill; STOP_MARKET protege toda la pata y el trailing agregado protege el remanente. Modo de posición: HEDGE ({'LONG' if signal['direction']=='BUY' else 'SHORT'}). Entorno: {mode_label()}."]
    return "\n".join(lines)

# ---------------------------- Envío de órdenes -------------------------------
def submit_trade(signal: dict[str, Any], plan: dict[str, Any], setup: dict[str, Any], chat_id: int) -> dict[str, Any]:
    api = BinanceUSDM()
    if not api.credentials_ok():
        raise RuntimeError(f"Configura BINANCE_{api.mode}_API_KEY y BINANCE_{api.mode}_API_SECRET de Binance Futures {api.mode} en .env.")
    symbol = signal["symbol"]
    validate_exit_settings(plan, setup)
    rules = api.symbol_rules(symbol)
    # Enforzar Hedge Mode (dualSidePosition=true). El cambio es de cuenta:
    # Binance exige no tener órdenes ni posiciones abiertas en ningún símbolo.
    api.ensure_hedge_mode()
    position_side = api.position_side(signal["direction"])
    same_leg_pos = api.position(symbol, signal["direction"])
    if abs(float(same_leg_pos.get("positionAmt", 0) or 0)) > 0:
        raise RuntimeError(f"Ya existe una posición {position_side} abierta en {symbol}. El bot no mezclará dos operaciones de la misma pata.")
    open_orders_all = api.request("GET", "/fapi/v1/openOrders", {"symbol": symbol}, signed=True)
    open_orders = [o for o in open_orders_all if o.get("positionSide") == position_side]
    if open_orders:
        raise RuntimeError(f"Ya existen órdenes abiertas de la pata {position_side} en {symbol}. El bot no las mezclará con esta operación.")
    open_algo_all = api.request("GET", "/fapi/v1/openAlgoOrders", {"algoType": "CONDITIONAL", "symbol": symbol}, signed=True)
    open_algo = [o for o in open_algo_all if o.get("positionSide") == position_side]
    if open_algo:
        raise RuntimeError(f"Ya existen órdenes condicionales/trailing de la pata {position_side} en {symbol}. Revísalas antes de continuar.")

    profile = plan["profile"]
    all_symbol_positions = api.positions(symbol)
    opposite_side = "SHORT" if position_side == "LONG" else "LONG"
    opposite_pos = next((x for x in all_symbol_positions if x.get("positionSide") == opposite_side), {})
    opposite_qty = abs(float(opposite_pos.get("positionAmt", 0) or 0))
    opposite_orders = [o for o in open_orders_all if o.get("positionSide") == opposite_side]
    opposite_algos = [o for o in open_algo_all if o.get("positionSide") == opposite_side]
    if opposite_qty <= 0 and (opposite_orders or opposite_algos):
        raise RuntimeError(
            f"Hay órdenes pendientes de la pata opuesta {opposite_side} en {symbol}, pero no una posición que permita validar sus ajustes. "
            "No cambiaré el apalancamiento/margen compartido del símbolo mientras esas órdenes existan."
        )
    if opposite_qty > 0:
        # El leverage/margin type es configuración del símbolo compartida por
        # ambas patas. No modificarlo si ya existe la posición opuesta.
        live_lev = int(float(opposite_pos.get("leverage", profile["leverage"])))
        margin_raw = str(opposite_pos.get("marginType", profile["margin_type"])).upper()
        live_margin = "CROSSED" if margin_raw in {"CROSS", "CROSSED"} else "ISOLATED" if margin_raw == "ISOLATED" else margin_raw
        if live_lev != int(round(profile["leverage"])) or live_margin != str(profile["margin_type"]).upper():
            raise RuntimeError(
                f"Ya existe {opposite_side} en {symbol} con {live_lev}x/{live_margin}. "
                f"El perfil de esta señal requiere {int(round(profile['leverage']))}x/{profile['margin_type']}. "
                "No cambiaré ajustes compartidos mientras la pata opuesta esté abierta."
            )
    else:
        api.set_symbol_settings(symbol, profile["margin_type"], int(round(profile["leverage"])))
    account = api.account()
    if account.get("canTrade") is False:
        raise RuntimeError(f"La cuenta Binance {api.mode} no tiene habilitado el permiso de trading de Futures.")
    available = float(account.get("availableBalance", 0))
    required_margin = float(plan["total_margin"])
    if required_margin > available * 0.98:
        raise RuntimeError(f"Saldo disponible insuficiente en Binance {api.mode}. Requiere aprox. {required_margin:.4f} USDT; disponible {available:.4f} USDT.")

    entries = []
    for e in plan["entries"]:
        price = quantize_price(float(e["price"]), rules["tick"], "nearest")
        qty_excel = float(e["quantity_excel"])
        qty = floor_step(qty_excel, rules["step"])
        if qty < rules["min_qty"] or qty * price < rules["min_notional"]:
            raise RuntimeError(f"Entrada {e['number']} queda por debajo del mínimo de Binance ({rules['min_qty']} unidades / {rules['min_notional']} USDT). Aumenta capital o reduce entradas.")
        entries.append({"number": e["number"], "price": price, "quantity": qty, "weight": e["weight"],
                        "margin": e["margin"], "order_id": None, "client_order_id": None,
                        "processed_qty": 0.0, "processed_quote": 0.0, "status": "NOT_SUBMITTED"})

    # Prevalidar que el nominal total no supere el capital calculado y que el split de salida quepa.
    total_qty = sum(x["quantity"] for x in entries)
    if total_qty <= 0:
        raise RuntimeError("La cantidad total redondeada quedó en cero.")

    trade_id = hashlib.sha1(f"{symbol}-{time.time_ns()}".encode()).hexdigest()[:12]
    trade = {
        "id": trade_id, "environment": api.mode, "symbol": symbol, "display_symbol": signal["display_symbol"],
        "direction": signal["direction"], "position_side": position_side,
        "chat_id": chat_id, "status": "ACTIVE",
        "created_at": int(time.time()), "setup": copy.deepcopy(setup), "profile": profile,
        "rules": rules, "entries": entries, "exit_slices": [],
        "pending_exit_qty": 0.0, "pending_exit_quote": 0.0,
        "total_planned_qty": total_qty,
    }
    # Persistimos el borrador antes de enviar la primera orden para que un
    # reinicio durante la colocación pueda recuperar las órdenes por client ID.
    trade["status"] = "SUBMITTING"
    DATA.setdefault("trades", {})[trade_id] = trade
    save_data()
    side = "BUY" if signal["direction"] == "BUY" else "SELL"
    placed: list[int] = []
    try:
        for entry in entries:
            cid = f"JOHAE{trade_id[-7:]}{entry['number']}"[:36]
            entry["client_order_id"] = cid
            entry["status"] = "SUBMISSION_UNKNOWN"
            save_data()  # el client ID queda guardado antes del POST
            try:
                res = api.new_limit(symbol, side, entry["quantity"], entry["price"], cid, position_side)
            except Exception as send_exc:
                try:
                    res = api.order_by_client_id(symbol, cid)
                except Exception:
                    raise send_exc
            entry["order_id"] = int(res["orderId"])
            entry["status"] = res.get("status", "NEW")
            placed.append(entry["order_id"])
            save_data()
    except Exception as exc:
        # Nunca abandonar fills ejecutados sin seguimiento: cancela las entradas
        # abiertas que alcanzaron a enviarse, guarda sus IDs y deja al monitor
        # procesar toda cantidad que sí haya quedado ejecutada.
        trade["submission_warning"] = f"Colocación incompleta: {exc}"
        for entry in entries:
            if entry.get("order_id") and entry.get("status") not in {"FILLED", "CANCELED", "EXPIRED", "REJECTED"}:
                try:
                    api.cancel_order(symbol, int(entry["order_id"]))
                    entry["status"] = "CANCELED"
                except Exception:
                    LOG.exception("No se pudo cancelar entrada %s al abortar colocación", entry.get("order_id"))
            elif not entry.get("order_id") and not entry.get("client_order_id"):
                entry["status"] = "CANCELED_BEFORE_SUBMIT"
            elif not entry.get("order_id") and entry.get("client_order_id"):
                entry["status"] = "SUBMISSION_UNKNOWN"
        if not placed and not any(e.get("status") == "SUBMISSION_UNKNOWN" for e in entries):
            trade["status"] = "FAILED"
            save_data()
            raise RuntimeError(trade["submission_warning"])
    trade["status"] = "ACTIVE"
    save_data()
    return trade

def execution_summary(trade: dict[str, Any]) -> str:
    label = mode_label(trade.get("environment", "DEMO"))
    return (f"✅ ENTRADAS ENVIADAS A BINANCE USDⓈ-M FUTURES — {label}\n{trade['display_symbol']} | {'LONG' if trade['direction']=='BUY' else 'SHORT'}\n"
            f"Órdenes límite: {len(trade['entries'])}\nMargen estimado: {sum(x['margin'] for x in trade['entries']):.4f} USDT\n"
            f"Las LIMIT parciales cubren {trade['setup']['partial_pct']}% de cada bloque ejecutado; un único trailing agregado protege el resto.\n"
            f"STOP_MARKET según Excel: -{trade['profile']['stop_roi_pct']:g}% ROI / {trade['profile']['leverage']:g}x, pata {trade['position_side']}.\n"
            "Modo de posición: HEDGE. Mantén el programa encendido para supervisar fills y salidas.")

def binance_status() -> str:
    mode = configured_binance_mode()
    api = BinanceUSDM(mode)
    acc = api.account()
    usdt = next((a for a in acc.get("assets", []) if a.get("asset") == "USDT"), {})
    hedge = api.is_hedge_mode()
    return (f"🔌 BINANCE USDⓈ-M FUTURES — {mode_label(mode)}\nConexión autenticada: OK\n"
            f"Modo de posición: {'HEDGE' if hedge else 'ONE-WAY (el bot lo cambiará a HEDGE al ejecutar, si Binance lo permite)'}\n"
            f"Wallet USDT: {usdt.get('walletBalance', '0')}\nDisponible: {acc.get('availableBalance', '0')} USDT\n"
            + ("⚠️ Producción habilitada: las órdenes pueden usar fondos reales." if mode == "LIVE" else "No se envían órdenes reales desde este entorno."))

# ----------------------------- Monitor fills ---------------------------------
def monitor_trade_once(trade: dict[str, Any]) -> list[str]:
    # Cada operación sigue usando el entorno donde fue abierta, aunque .env cambie luego.
    api = BinanceUSDM(trade.get("environment", "DEMO"))
    notices: list[str] = []
    symbol = trade["symbol"]
    rules = trade["rules"]
    # Actualiza fills incrementales de cada entrada.
    for entry in trade["entries"]:
        if not entry.get("order_id") and entry.get("client_order_id"):
            try:
                recovered = api.order_by_client_id(symbol, entry["client_order_id"])
                entry["order_id"] = int(recovered["orderId"])
                entry["status"] = recovered.get("status", "NEW")
                save_data()
            except Exception:
                # Puede que el POST no alcanzara Binance o que la respuesta se
                # haya perdido; no reenviamos la entrada para no duplicarla.
                continue
        if not entry.get("order_id"):
            continue
        try:
            current = api.order(symbol, int(entry["order_id"]))
        except Exception as exc:
            LOG.warning("No se pudo consultar orden %s: %s", entry.get("order_id"), exc)
            continue
        entry["status"] = current.get("status", entry.get("status"))
        executed = float(current.get("executedQty", 0))
        cum_quote = float(current.get("cumQuote", current.get("cumQuoteQty", 0)) or 0)
        dq = executed - float(entry.get("processed_qty", 0))
        dquote = cum_quote - float(entry.get("processed_quote", 0))
        if dq > float(rules["step"]) / 2:
            entry["processed_qty"] = executed
            entry["processed_quote"] = cum_quote
            trade["pending_exit_qty"] = float(trade.get("pending_exit_qty", 0)) + dq
            trade["pending_exit_quote"] = float(trade.get("pending_exit_quote", 0)) + (
                dquote if dquote > 0 else dq * float(current.get("avgPrice", entry["price"])))
            notices.append(f"Fill detectado {trade['display_symbol']}: {fmt(dq,8)}")

    # Crear/reintentar salidas para cantidades ejecutadas. Si la respuesta de
    # Binance se pierde, process_fill_batch recupera por clientOrderId/clientAlgoId.
    pending_qty = float(trade.get("pending_exit_qty", 0))
    pending_quote = float(trade.get("pending_exit_quote", 0))
    if pending_qty > 0:
        try:
            ready, message, consumed_qty = process_fill_batch(api, trade, pending_qty, pending_quote, rules)
            if ready:
                # process_fill_batch ya descontó el bloque protegido y lo guardó
                # atómicamente con la creación de la salida; no descontar aquí.
                trade["pending_exit_warning_sent"] = False
                notices.append(message)
            elif message and not trade.get("pending_exit_warning_sent"):
                # Evita notificaciones cada 3s, pero deja la cantidad pendiente guardada.
                notices.append("⚠️ " + message)
                trade["pending_exit_warning_sent"] = True
        except Exception as exc:
            LOG.exception("No se pudo crear salidas para %s", symbol)
            if not trade.get("pending_exit_warning_sent"):
                notices.append(f"🚨 No se pudo proteger un fill de {symbol}: {exc}. Revisa la operación en Binance {trade.get('environment', 'DEMO')}.")
                trade["pending_exit_warning_sent"] = True

    # Protege la pata con el stop fijo del Excel; después mantiene el trailing
    # para el remanente parcial, ambos en la misma posición LONG/SHORT.
    stop_fired = False
    try:
        stop_notes, stop_fired = ensure_fixed_stop_loss(api, trade)
        notices.extend(stop_notes)
        trade["stop_loss_present"] = not stop_fired
        trade["stop_loss_warning_sent"] = False
    except Exception as exc:
        LOG.exception("No se pudo ajustar el Stop Loss de %s", symbol)
        if not trade.get("stop_loss_warning_sent"):
            notices.append(f"🚨 No se pudo actualizar el STOP LOSS de {symbol}: {exc}. Revisa Binance {trade.get('environment', 'DEMO')} inmediatamente.")
            trade["stop_loss_warning_sent"] = True

    if not stop_fired:
        try:
            trail_notes, trail_present = ensure_aggregated_trailing(api, trade)
            notices.extend(trail_notes)
            trade["trailing_present"] = trail_present
            trade["pending_exit_warning_sent"] = False
        except Exception as exc:
            LOG.exception("No se pudo ajustar el trailing agregado para %s", symbol)
            if not trade.get("trailing_warning_sent"):
                notices.append(f"🚨 No se pudo actualizar el trailing de {symbol}: {exc}. Revisa la posición en Binance {trade.get('environment', 'DEMO')}.")
                trade["trailing_warning_sent"] = True

    # Consulta exclusivamente la pata de esta operación; la posición opuesta
    # del mismo símbolo permanece independiente en Hedge Mode.
    try:
        pos = api.position(symbol, trade["direction"])
        position_qty = abs(float(pos.get("positionAmt", 0)))
    except Exception as exc:
        LOG.warning("Consulta posición %s: %s", symbol, exc)
        position_qty = -1
    entries_terminal = all(e.get("status") in {"FILLED", "CANCELED", "EXPIRED", "REJECTED", "CANCELED_BEFORE_SUBMIT"}
                           for e in trade["entries"])
    has_exit = (any(sl.get("tp_status") == "FILLED" for sl in trade.get("exit_slices", [])) or
                str(trade.get("trail_order", {}).get("status", "")).upper() in {"TRIGGERED", "FINISHED"} or
                str(trade.get("stop_loss_order", {}).get("status", "")).upper() in {"TRIGGERED", "FINISHED", "MARKET_CLOSE_UNKNOWN"})
    has_fill = any(float(e.get("processed_qty", 0)) > 0 for e in trade["entries"])
    if position_qty == 0 and trade.get("pending_exit_qty", 0) == 0 and (entries_terminal or has_exit or has_fill):
        if has_exit or has_fill:
            for entry in trade["entries"]:
                if entry.get("order_id") and entry.get("status") not in {"FILLED", "CANCELED", "EXPIRED", "REJECTED"}:
                    try:
                        api.cancel_order(symbol, int(entry["order_id"]))
                        entry["status"] = "CANCELED"
                    except Exception as exc:
                        LOG.warning("No se pudo cancelar recompra pendiente %s: %s", entry.get("order_id"), exc)
            # Cancela salidas hermanas que sigan abiertas cuando la posición ya cerró.
            for sl in trade.get("exit_slices", []):
                if sl.get("tp_order_id") and sl.get("tp_status") not in {"FILLED", "CANCELED", "EXPIRED", "REJECTED"}:
                    try: api.cancel_order(symbol, int(sl["tp_order_id"]))
                    except Exception: pass
                    sl["tp_status"] = "CANCELED"
            trail_state = trade.get("trail_order", {})
            if trail_state.get("algo_id") and trail_state.get("status") not in {"FINISHED", "CANCELED", "EXPIRED", "REJECTED"}:
                try: api.cancel_algo(int(trail_state["algo_id"]))
                except Exception: pass
                trail_state["status"] = "CANCELED"
            stop_state = trade.get("stop_loss_order", {})
            if stop_state.get("algo_id") and stop_state.get("status") not in {"FINISHED", "TRIGGERED", "CANCELED", "EXPIRED", "REJECTED"}:
                try: api.cancel_algo(int(stop_state["algo_id"]))
                except Exception: pass
                stop_state["status"] = "CANCELED"
        trade["status"] = "CLOSED"
        notices.append(f"✅ Operación {trade['display_symbol']} cerrada. Se verificó posición cero.")
    save_data()
    return notices

async def monitor_loop(app: Application) -> None:
    while True:
        try:
            active = [(tid, t) for tid, t in DATA.get("trades", {}).items() if t.get("status") == "ACTIVE"]
            for tid, trade in active:
                try:
                    notes = await asyncio.to_thread(monitor_trade_once, trade)
                    for note in notes:
                        await app.bot.send_message(int(trade["chat_id"]), note)
                except Exception:
                    LOG.exception("Error monitoreando operación %s", tid)
            await asyncio.sleep(POLL_SECONDS)
        except asyncio.CancelledError:
            raise
        except Exception:
            LOG.exception("Error general del monitor")
            await asyncio.sleep(POLL_SECONDS)

def status_text() -> str:
    trades = [t for t in DATA.get("trades", {}).values() if t.get("status") == "ACTIVE"]
    legacy = [t for t in DATA.get("trades", {}).values() if t.get("status") == "LEGACY_REVIEW"]
    if not trades and not legacy:
        return "📈 No hay operaciones activas administradas por el bot."
    out = ["📈 OPERACIONES ACTIVAS"]
    for t in legacy:
        out.append(f"🚨 REVISIÓN MANUAL: {t.get('display_symbol', t.get('symbol','?'))} — no supervisada por versión Hedge")
    for t in trades:
        filled = sum(float(e.get("processed_qty", 0)) for e in t["entries"])
        out.append(f"{t['display_symbol']} {'LONG' if t['direction']=='BUY' else 'SHORT'} | {t.get('environment','DEMO')} | fills {fmt(filled,8)} | órdenes {len(t['entries'])} | salidas {len(t['exit_slices'])}")
    return "\n".join(out)

# --------------------------------- Main --------------------------------------
async def post_init(app: Application) -> None:
    # Alertar sobre operaciones guardadas por versiones antiguas One-way.
    for trade in DATA.get("trades", {}).values():
        if trade.get("status") == "LEGACY_REVIEW" and not trade.get("legacy_review_warning_sent"):
            warning = (
                f"🚨 OPERACIÓN ANTIGUA REQUIERE REVISIÓN MANUAL: {trade.get('display_symbol', trade.get('symbol','?'))}. "
                "Fue creada por una versión anterior sin positionSide LONG/SHORT. Esta versión Hedge no la supervisará. "
                "Verifica y gestiona esa posición y sus órdenes directamente en Binance antes de abrir operaciones nuevas."
            )
            try:
                await app.bot.send_message(int(trade["chat_id"]), warning)
                trade["legacy_review_warning_sent"] = True
            except Exception:
                LOG.exception("No se pudo notificar operación antigua")
    # Recuperación tras un cierre inesperado durante el envío del grid.
    for trade in DATA.get("trades", {}).values():
        if trade.get("status") == "SUBMITTING":
            for entry in trade.get("entries", []):
                if entry.get("status") == "NOT_SUBMITTED":
                    entry["status"] = "CANCELED_BEFORE_SUBMIT"
                elif not entry.get("order_id") and entry.get("client_order_id"):
                    entry["status"] = "SUBMISSION_UNKNOWN"
            trade["status"] = "ACTIVE"
            trade["submission_warning"] = f"El proceso se reinició durante el envío. El monitor está recuperando client IDs; revisa Binance {trade.get('environment', 'DEMO')} antes de reintentar."
            try:
                await app.bot.send_message(int(trade["chat_id"]), "⚠️ " + trade["submission_warning"])
            except Exception:
                LOG.exception("No se pudo notificar recuperación de una operación")
    save_data()
    app.create_task(monitor_loop(app))

async def post_shutdown(app: Application) -> None:
    # Persistir estado al cerrar Telegram.
    save_data()

def main() -> None:
    load_env()
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    if not token or "PEGA_AQUI" in token:
        raise SystemExit("Falta TELEGRAM_BOT_TOKEN. Coloca tu token de BotFather en .env.")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")
    app = Application.builder().token(token).post_init(post_init).post_shutdown(post_shutdown).build()
    app.add_handler(CommandHandler("start", start_cmd))
    app.add_handler(CommandHandler("config", config_cmd))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("status", status_cmd))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler((filters.TEXT | filters.CAPTION) & ~filters.COMMAND, handle_text))
    mode = configured_binance_mode()
    if mode == "LIVE" and not live_trading_allowed():
        print("AVISO: BINANCE_ENV=LIVE pero ALLOW_LIVE_TRADING no es YES; no se enviarán órdenes.")
    print(f"Trading bot iniciado — Binance USDⓈ-M Futures {mode_label(mode)} — Hedge Mode")
    app.run_polling(drop_pending_updates=False)

async def status_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        await update.effective_message.reply_text("No autorizado.")
        return
    await update.effective_message.reply_text(status_text(), reply_markup=menu_kb())

if __name__ == "__main__":
    main()
