from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
)
from copy import deepcopy
from datetime import datetime
from io import BytesIO
import os
import re
import math
import time
import traceback
import json

import pandas as pd
import requests


# ============================================================
# FLASK APP
# ============================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "poltrada-route-secret-key-2026"
)


# ============================================================
# KONFIGURASI GUDANG
# ============================================================

GUDANG_NAMA = "Gudang Distribusi Indomaret"

GUDANG_ALAMAT = (
    "Jl. Raya Mengwi No.17, Br. Binong, "
    "Desa Werdi Bhuwana, Kecamatan Mengwi"
)

GUDANG_LAT_TETAP = -8.544778
GUDANG_LON_TETAP = 115.166306


# ============================================================
# KONFIGURASI DEFAULT
# ============================================================

DEFAULT_JUMLAH_KENDARAAN = 3
DEFAULT_KAPASITAS = 5000
DEFAULT_HARGA_BBM = 6800
DEFAULT_BIAYA_SUPIR = 25000

# 1 liter = 5 km
BBM_KM_PER_LITER = 5.0

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OSRM_URL = "https://router.project-osrm.org"

USER_AGENT = (
    "PoltradaRouteOptimization/1.0 "
    "(route optimization educational project)"
)


# ============================================================
# GLOBAL STATE
# ============================================================

ACTIVE = {
    "jumlah_kendaraan": DEFAULT_JUMLAH_KENDARAAN,
    "kapasitas": DEFAULT_KAPASITAS,
    "harga_bbm": DEFAULT_HARGA_BBM,
    "biaya_supir": DEFAULT_BIAYA_SUPIR,

    "gudang": {
        "nama": GUDANG_NAMA,
        "alamat": GUDANG_ALAMAT,
        "lat": GUDANG_LAT_TETAP,
        "lon": GUDANG_LON_TETAP,
        "source": "default",
        "display_name": GUDANG_ALAMAT,
    },

    "outlets": [],

    "matrix_distance": None,
    "matrix_time": None,
    "matrix_codes": [],

    "last_api_time": None,
    "last_result": None,
}

HISTORY = []
REDO_HISTORY = []

IMPORT_PREVIEW = []

# Cache geocoding di runtime/server.
# Disimpan juga ke /tmp agar tidak melakukan pencarian berulang selama
# container masih hidup. Tidak berisi data rahasia.
GEOCODE_CACHE_FILE = os.path.join(
    os.environ.get("TMPDIR", "/tmp"),
    "poltrada_geocode_cache.json"
)
GEOCODE_CACHE = {}


# ============================================================
# UTILITAS
# ============================================================

def now_text():
    return datetime.now().strftime(
        "%d-%m-%Y %H:%M:%S"
    )


def clean_text(value):
    if value is None:
        return ""

    # Nilai kosong dari Excel sering terbaca sebagai NaN/NaT.
    # Jangan ubah nilai tersebut menjadi teks "nan" atau "nat".
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize_header(value):
    text = clean_text(value)

    text = text.lower()

    text = (
        text.replace(" ", "_")
        .replace("-", "_")
        .replace(".", "_")
        .replace("/", "_")
    )

    text = re.sub(
        r"_+",
        "_",
        text
    )

    return text.strip("_")


def to_float(value, default=None):
    """Mengubah nilai menjadi angka dengan toleransi format Excel/Indonesia.

    Mendukung contoh:
    3410
    3410.5
    3.410
    3,410
    3.410,5
    3,410.5
    """
    if value is None:
        return default

    if isinstance(value, bool):
        return default

    if isinstance(value, (int, float)):
        try:
            if pd.isna(value):
                return default
        except Exception:
            pass
        return float(value)

    text = clean_text(value)
    if not text:
        return default

    # Hilangkan simbol mata uang/unit yang sering ikut terbaca dari Excel.
    text = text.replace("Rp", "").replace("IDR", "")
    text = re.sub(r"[^0-9,.-]", "", text)

    if not text or text in {"-", ".", ","}:
        return default

    # Tangani tanda negatif hanya di depan.
    negative = text.startswith("-")
    text = text.lstrip("-")
    if not text:
        return default

    # Format Indonesia: 3.410,5 -> 3410.5
    # Format internasional: 3,410.5 -> 3410.5
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        # 3,410 -> 3410; 3410,5 -> 3410.5
        if len(parts) == 2 and len(parts[1]) <= 2:
            text = parts[0].replace(".", "") + "." + parts[1]
        else:
            text = "".join(parts)
    elif "." in text:
        parts = text.split(".")
        # 3.410 -> 3410 (pemisah ribuan), 3410.5 -> 3410.5
        if len(parts) > 2:
            text = "".join(parts)
        elif len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) <= 3:
            text = "".join(parts)

    try:
        result = float(text)
        return -result if negative else result
    except Exception:
        return default


def safe_float(value, default=0.0):
    """Versi aman to_float untuk perhitungan dan tampilan."""
    result = to_float(value, default)
    if result is None:
        return default
    return result


def normalize_matrix_code(value):
    if isinstance(value, dict):
        value = (
            value.get("kode")
            or value.get("code")
            or value.get("KODE")
            or value.get("CODE")
            or value.get("id")
            or value.get("ID")
            or ""
        )

    if isinstance(value, dict):
        value = ""

    return str(value).strip().upper()


# ============================================================
# HISTORY
# ============================================================

def snapshot_state():
    return {
        "outlets": deepcopy(
            ACTIVE["outlets"]
        ),

        "matrix_distance": deepcopy(
            ACTIVE["matrix_distance"]
        ),

        "matrix_time": deepcopy(
            ACTIVE["matrix_time"]
        ),

        "matrix_codes": deepcopy(
            ACTIVE["matrix_codes"]
        ),

        "last_api_time": ACTIVE[
            "last_api_time"
        ],

        "last_result": deepcopy(
            ACTIVE["last_result"]
        ),

        "jumlah_kendaraan": ACTIVE[
            "jumlah_kendaraan"
        ],

        "kapasitas": ACTIVE[
            "kapasitas"
        ],

        "harga_bbm": ACTIVE[
            "harga_bbm"
        ],

        "biaya_supir": ACTIVE[
            "biaya_supir"
        ],

        "gudang": deepcopy(
            ACTIVE["gudang"]
        ),
    }


def restore_state(state):
    ACTIVE["outlets"] = deepcopy(
        state["outlets"]
    )

    ACTIVE["matrix_distance"] = deepcopy(
        state["matrix_distance"]
    )

    ACTIVE["matrix_time"] = deepcopy(
        state["matrix_time"]
    )

    ACTIVE["matrix_codes"] = deepcopy(
        state["matrix_codes"]
    )

    ACTIVE["last_api_time"] = state[
        "last_api_time"
    ]

    ACTIVE["last_result"] = deepcopy(
        state["last_result"]
    )

    ACTIVE["jumlah_kendaraan"] = state[
        "jumlah_kendaraan"
    ]

    ACTIVE["kapasitas"] = state[
        "kapasitas"
    ]

    ACTIVE["harga_bbm"] = state[
        "harga_bbm"
    ]

    ACTIVE["biaya_supir"] = state[
        "biaya_supir"
    ]

    ACTIVE["gudang"] = deepcopy(
        state.get(
            "gudang",
            ACTIVE["gudang"]
        )
    )


def save_history():
    HISTORY.append(
        snapshot_state()
    )

    if len(HISTORY) > 50:
        HISTORY.pop(0)

    REDO_HISTORY.clear()


# ============================================================
# NORMALISASI DATA OUTLET
# ============================================================

HEADER_ALIASES = {
    "kode": [
        "kode",
        "kode_outlet",
        "kode_toko",
        "code",
        "id",
        "outlet",
    ],

    "nama": [
        "nama",
        "nama_outlet",
        "nama_toko",
        "name",
        "toko",
    ],

    "alamat": [
        "alamat",
        "alamat_outlet",
        "alamat_toko",
        "address",
        "lokasi",
    ],

    "permintaan": [
        "permintaan",
        "demand",
        "kebutuhan",
        "jumlah",
        "muatan",
        "qty",
        "quantity",
    ],

    # Koordinat opsional. Jika tersedia di Excel, sistem tidak perlu
    # melakukan geocoding untuk outlet tersebut.
    "lat": [
        "lat",
        "latitude",
        "lintang",
        "latitude_outlet",
        "lat_outlet",
        "koordinat_lat",
    ],

    "lon": [
        "lon",
        "lng",
        "longitude",
        "bujur",
        "longitude_outlet",
        "lon_outlet",
        "koordinat_lon",
    ],
}


def canonicalize_columns(df):
    rename_map = {}
    normalized = {}

    for col in df.columns:
        normalized[
            normalize_header(col)
        ] = col

    for canonical, aliases in HEADER_ALIASES.items():

        found = None

        for alias in aliases:

            alias_norm = normalize_header(
                alias
            )

            if alias_norm in normalized:

                found = normalized[
                    alias_norm
                ]

                break

        if found is not None:

            rename_map[
                found
            ] = canonical

    return df.rename(
        columns=rename_map
    )


def validate_outlet_dataframe(df):

    errors = []

    df = canonicalize_columns(df)

    required = [
        "kode",
        "nama",
        "alamat",
        "permintaan",
    ]

    missing = [
        x
        for x in required
        if x not in df.columns
    ]

    if missing:

        errors.append(
            "Kolom wajib belum lengkap: "
            + ", ".join(missing)
        )

        return df, errors

    records = []
    seen_codes = set()

    for idx, row in df.iterrows():

        nomor = idx + 2

        kode = clean_text(
            row.get("kode")
        )

        nama = clean_text(
            row.get("nama")
        )

        alamat = clean_text(
            row.get("alamat")
        )

        raw_permintaan = row.get("permintaan")
        permintaan = to_float(
            raw_permintaan
        )

        # Lewati baris Excel yang benar-benar kosong.
        # Sebelumnya baris kosong terbaca sebagai NaN lalu dianggap
        # sebagai data outlet sehingga muncul error "permintaan tidak valid".
        if not kode and not nama and not alamat and permintaan is None:
            continue

        if not kode:
            errors.append(
                f"Baris {nomor}: kode outlet kosong."
            )
            continue

        kode_upper = normalize_matrix_code(kode)

        if kode_upper in seen_codes:
            errors.append(
                f"Baris {nomor}: kode outlet "
                f"{kode} duplikat."
            )
            continue

        seen_codes.add(kode_upper)

        if not nama:
            errors.append(
                f"Baris {nomor}: nama outlet kosong."
            )
            continue

        if not alamat:
            errors.append(
                f"Baris {nomor}: alamat outlet kosong."
            )
            continue

        if permintaan is None:
            raw_text = clean_text(raw_permintaan)
            errors.append(
                f"Baris {nomor}: permintaan tidak valid (nilai sel: {raw_text or 'kosong'})."
            )
            continue

        if permintaan <= 0:
            errors.append(
                f"Baris {nomor}: permintaan "
                "harus lebih besar dari 0."
            )
            continue

        lat_value = to_float(row.get("lat"), None) if "lat" in df.columns else None
        lon_value = to_float(row.get("lon"), None) if "lon" in df.columns else None

        if lat_value is not None and not (-90 <= lat_value <= 90):
            errors.append(
                f"Baris {nomor}: latitude tidak valid."
            )
            continue

        if lon_value is not None and not (-180 <= lon_value <= 180):
            errors.append(
                f"Baris {nomor}: longitude tidak valid."
            )
            continue

        # Jika hanya salah satu koordinat yang diisi, jangan diam-diam
        # menggunakan koordinat yang tidak lengkap.
        if (lat_value is None) != (lon_value is None):
            errors.append(
                f"Baris {nomor}: latitude dan longitude harus diisi berpasangan."
            )
            continue

        records.append({
            "kode": kode,
            "nama": nama,
            "alamat": alamat,
            "permintaan": float(permintaan),
            "lat": lat_value,
            "lon": lon_value,
        })

    result = pd.DataFrame(records)

    return result, errors


# ============================================================
# EXCEL IMPORT
# ============================================================

def read_retailer_excel(file_storage):

    filename = clean_text(
        file_storage.filename
    )

    if not filename:
        raise ValueError(
            "File Excel belum dipilih."
        )

    extension = os.path.splitext(
        filename
    )[1].lower()

    if extension not in [
        ".xlsx",
        ".xlsm",
    ]:
        raise ValueError(
            "File harus berformat "
            ".xlsx atau .xlsm."
        )

    try:
        df = pd.read_excel(
            file_storage
        )

    except Exception as e:
        raise ValueError(
            "File Excel tidak dapat dibaca: "
            + str(e)
        )

    if df.empty:
        raise ValueError(
            "File Excel tidak memiliki data."
        )

    df, errors = validate_outlet_dataframe(
        df
    )

    if errors:
        raise ValueError(
            "\n".join(errors)
        )

    if df.empty:
        raise ValueError(
            "Tidak ada data outlet yang valid."
        )

    return df.to_dict(
        orient="records"
    )


# ============================================================
# TEMPLATE CONTEXT
# ============================================================

@app.context_processor
def inject_global_data():

    result = ACTIVE[
        "last_result"
    ]

    ACTIVE["outlet"] = ACTIVE["outlets"]

    legacy_result = None

    if result:

        optimized = result.get(
            "optimized",
            {}
        )

        initial = result.get(
            "initial",
            {}
        )

        initial_distance = safe_float(
            initial.get(
                "total_distance"
            ),
            optimized.get(
                "total_distance",
                0
            )
        )

        optimized_distance = safe_float(
            optimized.get(
                "total_distance"
            ),
            0
        )

        persen = 0.0

        if initial_distance > 0:

            persen = (
                (
                    initial_distance
                    -
                    optimized_distance
                )
                /
                initial_distance
                *
                100.0
            )

        legacy_result = {

            "jarak_awal":
                initial_distance,

            "jarak_optimasi":
                optimized_distance,

            "persen":
                persen,

            "total_muatan":
                optimized.get(
                    "total_load",
                    0
                ),

            "total_waktu":
                optimized.get(
                    "total_time",
                    0
                ),

            "total_biaya":
                optimized.get(
                    "total_cost",
                    0
                ),

            "timestamp":
                result.get(
                    "meta",
                    {}
                ).get(
                    "created_at",
                    ""
                ),

            "api_status":
                "OSRM + Nominatim",

            "metode":
                "VRP + Nearest Neighbor",

            "rute": []
        }

        for rr in optimized.get(
            "routes",
            []
        ):

            item = dict(rr)

            item["jenis"] = "CDD 5 Ton"

            item["biaya"] = rr.get(
                "total_biaya",
                0
            )

            legacy_result[
                "rute"
            ].append(
                item
            )

    ACTIVE["hasil"] = legacy_result

    return {

        "active":
            ACTIVE,

        "result":
            result,

        "outlets":
            ACTIVE[
                "outlets"
            ],

        "gudang":
            ACTIVE[
                "gudang"
            ],

        "jumlah_kendaraan":
            ACTIVE[
                "jumlah_kendaraan"
            ],

        "kapasitas":
            ACTIVE[
                "kapasitas"
            ],

        "harga_bbm":
            ACTIVE[
                "harga_bbm"
            ],

        "biaya_supir":
            ACTIVE[
                "biaya_supir"
            ],

        "now":
            now_text(),

        "can_undo":
            bool(HISTORY),

        "can_redo":
            bool(REDO_HISTORY),

        "auto_result":
            bool(result),
    }


# ============================================================
# INDEX
# ============================================================

@app.route("/")
def index():

    return render_template(

        "index.html",

        outlets=ACTIVE[
            "outlets"
        ],

        result=ACTIVE[
            "last_result"
        ],

        gudang=ACTIVE[
            "gudang"
        ],

        import_preview=
            IMPORT_PREVIEW,

        import_total=
            len(
                IMPORT_PREVIEW
            ),

        import_demand=sum(
            safe_float(
                x.get(
                    "permintaan"
                )
            )
            for x in IMPORT_PREVIEW
        ),

        jumlah_kendaraan=
            ACTIVE[
                "jumlah_kendaraan"
            ],

        kapasitas=
            ACTIVE[
                "kapasitas"
            ],

        harga_bbm=
            ACTIVE[
                "harga_bbm"
            ],

        biaya_supir=
            ACTIVE[
                "biaya_supir"
            ],
    )


# ============================================================
# IMPORT EXCEL
# ============================================================

@app.post("/import-excel")
def import_excel():

    global IMPORT_PREVIEW

    file_excel = request.files.get(
        "file_excel"
    )

    if not file_excel:

        flash(
            "File Excel belum dipilih.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    try:

        records = read_retailer_excel(
            file_excel
        )

        IMPORT_PREVIEW = records

        total = len(
            records
        )

        demand = sum(
            safe_float(
                x.get(
                    "permintaan"
                )
            )
            for x in records
        )

        flash(
            f"File berhasil dibaca. "
            f"{total} outlet siap dikonfirmasi.",
            "success"
        )

        return render_template(

            "index.html",

            outlets=ACTIVE[
                "outlets"
            ],

            result=ACTIVE[
                "last_result"
            ],

            gudang=ACTIVE[
                "gudang"
            ],

            import_preview=
                IMPORT_PREVIEW,

            import_total=
                total,

            import_demand=
                demand,

            jumlah_kendaraan=
                ACTIVE[
                    "jumlah_kendaraan"
                ],

            kapasitas=
                ACTIVE[
                    "kapasitas"
                ],

            harga_bbm=
                ACTIVE[
                    "harga_bbm"
                ],

            biaya_supir=
                ACTIVE[
                    "biaya_supir"
                ],
        )

    except Exception as e:

        flash(
            f"Import gagal: {e}",
            "error"
        )

        return redirect(
            url_for("index")
        )


# ============================================================
# KONFIRMASI IMPORT
# ============================================================

@app.post("/konfirmasi-import")
def konfirmasi_import():

    global IMPORT_PREVIEW

    if not IMPORT_PREVIEW:

        flash(
            "Belum ada data import "
            "yang dapat dikonfirmasi.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    save_history()

    new_active_outlets = []

    try:

        for item in IMPORT_PREVIEW:

            kode = clean_text(
                item.get("kode")
            )

            nama = clean_text(
                item.get("nama")
            )

            alamat = clean_text(
                item.get("alamat")
            )

            permintaan = to_float(
                item.get("permintaan")
            )

            if not kode:
                raise ValueError(
                    "Terdapat outlet tanpa kode."
                )

            if not nama:
                raise ValueError(
                    f"Nama outlet "
                    f"{kode} kosong."
                )

            if not alamat:
                raise ValueError(
                    f"Alamat outlet "
                    f"{kode} kosong."
                )

            if (
                permintaan is None
                or permintaan <= 0
            ):

                raise ValueError(
                    f"Permintaan outlet "
                    f"{kode} tidak valid."
                )

            new_active_outlets.append({

                "kode":
                    kode,

                "nama":
                    nama,

                "alamat":
                    alamat,

                "permintaan":
                    float(
                        permintaan
                    ),

                "lat":
                    item.get(
                        "lat"
                    ),

                "lon":
                    item.get(
                        "lon"
                    ),
            })

        ACTIVE[
            "outlets"
        ] = new_active_outlets

        IMPORT_PREVIEW = []

        ACTIVE[
            "matrix_distance"
        ] = None

        ACTIVE[
            "matrix_time"
        ] = None

        ACTIVE[
            "matrix_codes"
        ] = []

        ACTIVE[
            "last_api_time"
        ] = None

        ACTIVE[
            "last_result"
        ] = None

        flash(
            "Import berhasil dikonfirmasi. "
            f"{len(ACTIVE['outlets'])} outlet "
            "menjadi data aktif.",
            "success"
        )

    except Exception as e:

        flash(
            f"Konfirmasi import gagal: {e}",
            "error"
        )

    return redirect(
        url_for("index")
    )


# ============================================================
# BATALKAN IMPORT
# ============================================================

@app.post("/batalkan-import")
def batalkan_import():

    global IMPORT_PREVIEW

    IMPORT_PREVIEW = []

    flash(
        "Preview import dibatalkan.",
        "info"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# GEOCODING
# ============================================================

def load_geocode_cache():
    """Membaca cache geocoding dari file sementara jika tersedia."""
    global GEOCODE_CACHE

    if GEOCODE_CACHE:
        return GEOCODE_CACHE

    try:
        if os.path.exists(GEOCODE_CACHE_FILE):
            with open(
                GEOCODE_CACHE_FILE,
                "r",
                encoding="utf-8"
            ) as f:
                data = json.load(f)

            if isinstance(data, dict):
                GEOCODE_CACHE = data
    except Exception:
        GEOCODE_CACHE = {}

    return GEOCODE_CACHE


def save_geocode_cache():
    """Menyimpan cache geocoding secara aman."""
    try:
        with open(
            GEOCODE_CACHE_FILE,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                GEOCODE_CACHE,
                f,
                ensure_ascii=False,
                indent=2
            )
    except Exception:
        # Cache hanya optimasi; kegagalan menulis cache tidak boleh
        # menghentikan perhitungan rute.
        pass


def geocode_cache_key(outlet):
    nama = clean_text(outlet.get("nama")).lower()
    alamat = clean_text(outlet.get("alamat")).lower()
    return f"{nama}|{alamat}"


def get_cached_geocode(outlet):
    cache = load_geocode_cache()
    key = geocode_cache_key(outlet)
    value = cache.get(key)

    if not isinstance(value, dict):
        return None

    lat = to_float(value.get("lat"), None)
    lon = to_float(value.get("lon"), None)

    if lat is None or lon is None:
        return None

    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return None

    return {
        "lat": lat,
        "lon": lon,
        "source": "cache",
        "display_name": value.get(
            "display_name",
            outlet.get("alamat") or outlet.get("nama")
        ),
    }


def set_cached_geocode(outlet, result):
    key = geocode_cache_key(outlet)

    GEOCODE_CACHE[key] = {
        "lat": result["lat"],
        "lon": result["lon"],
        "source": result.get("source", "geocoder"),
        "display_name": result.get(
            "display_name",
            outlet.get("alamat") or outlet.get("nama")
        ),
        "cached_at": now_text(),
    }

    save_geocode_cache()


def extract_plus_code(text):
    """Mengambil Plus Code sederhana dari alamat, jika ada."""
    text = clean_text(text).upper()

    match = re.search(
        r"(?<![A-Z0-9])([23456789CFGHJMPQRVWX]{4,8}\+[23456789CFGHJMPQRVWX]{2,})(?![A-Z0-9])",
        text
    )

    return match.group(1) if match else ""


def add_query_unique(container, value):
    value = clean_text(value)

    if value and value.lower() not in {
        x.lower() for x in container
    }:
        container.append(value)


def nominatim_search(
    query
):

    headers = {
        "User-Agent":
            USER_AGENT
    }

    params = {

        "q":
            query,

        "format":
            "json",

        "limit":
            1,

        "countrycodes":
            "id",
    }

    response = requests.get(

        NOMINATIM_URL,

        params=params,

        headers=headers,

        timeout=20
    )

    response.raise_for_status()

    data = response.json()

    if not data:
        return None

    return {

        "lat":
            float(
                data[0]["lat"]
            ),

        "lon":
            float(
                data[0]["lon"]
            ),

        "display_name":
            data[0].get(
                "display_name",
                query
            ),
    }


def generate_geocode_queries(
    outlet
):
    """Membuat variasi query yang lebih tahan terhadap alamat bisnis."""
    nama = clean_text(outlet.get("nama"))
    alamat = clean_text(outlet.get("alamat"))

    queries = []

    def add(value):
        add_query_unique(queries, value)

    # Query paling spesifik.
    if nama and alamat:
        add(f"{nama}, {alamat}")
        add(f"{nama}, {alamat}, Bali, Indonesia")
        add(f"{nama}, {alamat}, Indonesia")

    if alamat:
        add(f"{alamat}, Bali, Indonesia")
        add(f"{alamat}, Indonesia")

    # Plus Code sering lebih berguna daripada nama toko.
    plus_code = extract_plus_code(alamat)
    if plus_code:
        add(f"{plus_code}, Baturiti, Tabanan, Bali, Indonesia")
        add(f"{plus_code}, Tabanan, Bali, Indonesia")
        if nama:
            add(f"{nama}, {plus_code}, Baturiti, Bali, Indonesia")

    # Hilangkan kode pos.
    alamat_ringkas = re.sub(
        r"\b\d{5,6}\b",
        "",
        alamat
    )
    alamat_ringkas = re.sub(
        r"\s+",
        " ",
        alamat_ringkas
    ).strip(" ,.-")

    if nama and alamat_ringkas:
        add(f"{nama}, {alamat_ringkas}, Bali, Indonesia")
        add(f"{alamat_ringkas}, Bali, Indonesia")

    # Hilangkan nomor rumah/bangunan.
    alamat_tanpa_nomor = re.sub(
        r"\bNo\.?\s*[\w./-]+",
        "",
        alamat_ringkas,
        flags=re.I
    )
    alamat_tanpa_nomor = re.sub(
        r"\s+",
        " ",
        alamat_tanpa_nomor
    ).strip(" ,.-")

    if nama and alamat_tanpa_nomor:
        add(f"{nama}, {alamat_tanpa_nomor}, Bali, Indonesia")
        add(f"{alamat_tanpa_nomor}, Bali, Indonesia")

    # Nama saja sebagai fallback.
    if nama:
        add(f"{nama}, Bali, Indonesia")
        add(f"{nama}, Indonesia")

    return queries

def geocode_gudang(
    nama,
    alamat,
    lat=None,
    lon=None
):
    """Menentukan koordinat gudang secara dinamis."""
    lat_value = to_float(lat, None)
    lon_value = to_float(lon, None)

    if lat_value is not None and lon_value is not None:
        if -90 <= lat_value <= 90 and -180 <= lon_value <= 180:
            return {
                "lat": lat_value,
                "lon": lon_value,
                "source": "manual",
                "display_name": alamat or nama,
            }
        raise ValueError(
            "Koordinat gudang tidak valid. "
            "Latitude harus -90 s.d. 90 dan longitude -180 s.d. 180."
        )

    nama = clean_text(nama)
    alamat = clean_text(alamat)

    if not alamat and not nama:
        raise ValueError("Nama atau alamat gudang wajib diisi.")

    queries = []
    if alamat:
        queries.extend([
            alamat,
            f"{alamat}, Bali, Indonesia",
            f"{alamat}, Indonesia",
        ])
    if nama and alamat:
        queries.extend([
            f"{nama}, {alamat}",
            f"{nama}, {alamat}, Bali, Indonesia",
        ])
    elif nama:
        queries.extend([
            f"{nama}, Bali, Indonesia",
            f"{nama}, Indonesia",
        ])

    unique_queries = []
    seen = set()
    for query in queries:
        key = query.strip().lower()
        if key and key not in seen:
            seen.add(key)
            unique_queries.append(query.strip())

    last_error = None
    for query in unique_queries:
        for provider in ("nominatim", "arcgis", "photon"):
            try:
                result = geocode_public_provider(query, provider)
                if result:
                    return {
                        "lat": result["lat"],
                        "lon": result["lon"],
                        "source": provider,
                        "display_name": result.get("display_name", query),
                    }
            except Exception as e:
                last_error = e
        time.sleep(0.5)

    detail = f" Detail terakhir: {last_error}" if last_error else ""
    raise ValueError(
        "Koordinat gudang tidak ditemukan untuk "
        f"'{nama or '-'}', '{alamat or '-'}'. "
        "Gunakan alamat yang lebih lengkap atau isi latitude dan longitude manual."
        + detail
    )


def geocode_public_provider(query, provider):
    """Geocoder publik dengan format hasil seragam."""
    query = clean_text(query)
    if not query:
        return None

    headers = {
        "User-Agent": "PoltradaRouteOptimization/1.0",
        "Accept-Language": "id,en",
    }

    if provider == "nominatim":
        response = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": query,
                "format": "jsonv2",
                "limit": 3,
                "countrycodes": "id",
                "addressdetails": 1,
            },
            headers=headers,
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()
        if not data:
            return None
        # Pilih hasil pertama yang benar-benar memiliki koordinat.
        for item in data:
            if item.get("lat") and item.get("lon"):
                return {
                    "lat": float(item["lat"]),
                    "lon": float(item["lon"]),
                    "display_name": item.get("display_name", query),
                }
        return None

    if provider == "arcgis":
        response = requests.get(
            "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates",
            params={
                "SingleLine": query,
                "f": "json",
                "maxLocations": 5,
                "outFields": "*",
                "forStorage": "false",
            },
            headers=headers,
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()
        candidates = data.get("candidates") or []
        for item in candidates:
            location = item.get("location") or {}
            score = float(item.get("score", 0) or 0)
            if "x" in location and "y" in location and score >= 70:
                return {
                    "lat": float(location["y"]),
                    "lon": float(location["x"]),
                    "display_name": item.get("address", query),
                }
        return None

    if provider == "photon":
        response = requests.get(
            "https://photon.komoot.io/api/",
            params={"q": query, "limit": 5},
            headers=headers,
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()
        features = data.get("features") or []
        for item in features:
            geometry = item.get("geometry") or {}
            coords = geometry.get("coordinates") or []
            if len(coords) >= 2:
                props = item.get("properties") or {}
                display_parts = [
                    props.get("name"),
                    props.get("street"),
                    props.get("city"),
                    props.get("state"),
                    props.get("country"),
                ]
                return {
                    "lat": float(coords[1]),
                    "lon": float(coords[0]),
                    "display_name": ", ".join(
                        str(x).strip() for x in display_parts if x
                    ) or query,
                }
        return None

    raise ValueError(f"Provider geocoding tidak dikenal: {provider}")


def geocode_outlet(outlet):
    """Geocode outlet dengan koordinat manual, cache, lalu multi-provider."""
    lat = to_float(outlet.get("lat"), None)
    lon = to_float(outlet.get("lon"), None)

    # 1. Koordinat dari Excel/manual selalu diprioritaskan.
    if lat is not None and lon is not None:
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return {
                "lat": lat,
                "lon": lon,
                "source": "existing",
                "display_name": outlet.get("alamat") or outlet.get("nama"),
            }

    # 2. Gunakan cache agar outlet yang sama tidak terus meminta API.
    cached = get_cached_geocode(outlet)
    if cached:
        return cached

    queries = generate_geocode_queries(outlet)
    last_error = None

    # Alamat bisnis di Indonesia kadang tidak ditemukan provider pertama.
    providers = ("nominatim", "arcgis", "photon")

    for query in queries:
        for provider in providers:
            try:
                result = geocode_public_provider(
                    query,
                    provider
                )

                if result:
                    final_result = {
                        "lat": result["lat"],
                        "lon": result["lon"],
                        "source": provider,
                        "display_name": result.get(
                            "display_name",
                            query
                        ),
                    }

                    set_cached_geocode(
                        outlet,
                        final_result
                    )

                    return final_result

            except Exception as e:
                last_error = e

        # Jangan melakukan burst request ke provider publik.
        time.sleep(0.8)

    nama = clean_text(outlet.get("nama"))
    alamat = clean_text(outlet.get("alamat"))

    detail = (
        f" Detail terakhir: {last_error}"
        if last_error
        else ""
    )

    raise ValueError(
        "Koordinat tidak ditemukan untuk: "
        f"{nama}, {alamat}. "
        "Sistem sudah mencoba beberapa variasi alamat, "
        "Plus Code (jika tersedia), Nominatim, ArcGIS, dan Photon. "
        "Untuk outlet yang belum terdaftar di geocoder, "
        "isi kolom Latitude dan Longitude pada Excel."
        + detail
    )

def geocode_all_outlets(
    outlets
):

    result = []

    for outlet in outlets:

        item = deepcopy(
            outlet
        )

        geo = geocode_outlet(
            item
        )

        item["lat"] = geo[
            "lat"
        ]

        item["lon"] = geo[
            "lon"
        ]

        result.append(
            item
        )

    return result


# ============================================================
# OSRM ROUTING
# ============================================================

def build_osrm_coordinates(
    outlets
):

    gudang = ACTIVE.get("gudang", {})

    warehouse_lat = to_float(
        gudang.get("lat"),
        None
    )
    warehouse_lon = to_float(
        gudang.get("lon"),
        None
    )

    if warehouse_lat is None or warehouse_lon is None:
        raise ValueError(
            "Koordinat gudang belum tersedia."
        )

    coordinates = [

        (
            warehouse_lon,
            warehouse_lat
        )
    ]

    for outlet in outlets:

        coordinates.append(

            (
                float(
                    outlet["lon"]
                ),

                float(
                    outlet["lat"]
                )
            )
        )

    return coordinates


def calculate_api_matrix(
    outlets
):

    if not outlets:

        raise ValueError(
            "Belum ada data outlet."
        )

    geocoded_outlets = (
        geocode_all_outlets(
            outlets
        )
    )

    coordinates = (
        build_osrm_coordinates(
            geocoded_outlets
        )
    )

    coordinate_text = ";".join(

        f"{lon},{lat}"

        for lon, lat
        in coordinates
    )

    url = (
        f"{OSRM_URL}/table/v1/driving/"
        f"{coordinate_text}"
    )

    params = {
        "annotations":
            "distance,duration"
    }

    response = requests.get(

        url,

        params=params,

        headers={
            "User-Agent":
                USER_AGENT
        },

        timeout=120
    )

    response.raise_for_status()

    data = response.json()

    if data.get(
        "code"
    ) != "Ok":

        raise ValueError(

            "OSRM gagal menghitung "
            "matrix jarak dan waktu: "
            +
            str(
                data.get(
                    "message",
                    "Unknown error"
                )
            )
        )

    distances_m = data.get(
        "distances"
    )

    durations_s = data.get(
        "durations"
    )

    if not distances_m:

        raise ValueError(
            "Matrix jarak kosong."
        )

    if not durations_s:

        raise ValueError(
            "Matrix waktu kosong."
        )

    matrix_codes = [
        "GUDANG"
    ]

    matrix_codes.extend(

        normalize_matrix_code(
            x.get("kode")
        )

        for x
        in geocoded_outlets
    )

    ACTIVE[
        "matrix_distance"
    ] = distances_m

    ACTIVE[
        "matrix_time"
    ] = durations_s

    ACTIVE[
        "matrix_codes"
    ] = matrix_codes

    ACTIVE[
        "last_api_time"
    ] = now_text()

    outlet_map = {

        normalize_matrix_code(
            x.get("kode")
        ):
            x

        for x
        in geocoded_outlets
    }

    for outlet in ACTIVE[
        "outlets"
    ]:

        code = normalize_matrix_code(
            outlet.get(
                "kode"
            )
        )

        if code in outlet_map:

            outlet["lat"] = (
                outlet_map[
                    code
                ]["lat"]
            )

            outlet["lon"] = (
                outlet_map[
                    code
                ]["lon"]
            )

    return (

        distances_m,

        durations_s,

        matrix_codes
    )


# ============================================================
# MATRIX INDEX
# ============================================================

def matrix_index():

    result = {}

    codes = ACTIVE[
        "matrix_codes"
    ]

    for i, code in enumerate(
        codes
    ):

        normalized = (
            normalize_matrix_code(
                code
            )
        )

        if normalized:

            result[
                normalized
            ] = i

    return result


def get_matrix_distance(
    from_code,
    to_code
):

    distances = ACTIVE[
        "matrix_distance"
    ]

    if distances is None:

        raise ValueError(
            "Matrix jarak belum tersedia."
        )

    index = matrix_index()

    a = index.get(
        normalize_matrix_code(
            from_code
        )
    )

    b = index.get(
        normalize_matrix_code(
            to_code
        )
    )

    if a is None or b is None:

        raise ValueError(

            f"Kode matrix tidak ditemukan: "
            f"{from_code} -> {to_code}"
        )

    value = distances[
        a
    ][
        b
    ]

    if value is None:

        raise ValueError(

            f"Jarak tidak tersedia: "
            f"{from_code} -> {to_code}"
        )

    return float(
        value
    ) / 1000.0


def get_matrix_time(
    from_code,
    to_code
):

    durations = ACTIVE[
        "matrix_time"
    ]

    if durations is None:

        raise ValueError(
            "Matrix waktu belum tersedia."
        )

    index = matrix_index()

    a = index.get(
        normalize_matrix_code(
            from_code
        )
    )

    b = index.get(
        normalize_matrix_code(
            to_code
        )
    )

    if a is None or b is None:

        raise ValueError(

            f"Kode matrix tidak ditemukan: "
            f"{from_code} -> {to_code}"
        )

    value = durations[
        a
    ][
        b
    ]

    if value is None:

        raise ValueError(

            f"Waktu tidak tersedia: "
            f"{from_code} -> {to_code}"
        )

    return float(
        value
    ) / 60.0


# ============================================================
# ROUTE DISTANCE & TIME
# ============================================================

def calculate_route_distance(
    route
):

    if not route:

        return 0.0

    total = 0.0

    previous = "GUDANG"

    for code in route:

        total += get_matrix_distance(

            previous,

            code
        )

        previous = code

    total += get_matrix_distance(

        previous,

        "GUDANG"
    )

    return total


def calculate_route_time(
    route
):

    if not route:

        return 0.0

    total = 0.0

    previous = "GUDANG"

    for code in route:

        total += get_matrix_time(

            previous,

            code
        )

        previous = code

    total += get_matrix_time(

        previous,

        "GUDANG"
    )

    return total


# ============================================================
# NEAREST NEIGHBOR
# ============================================================

def nearest_neighbor(
    group
):

    if not group:

        return []

    remaining = {}

    for outlet in group:

        if isinstance(
            outlet,
            dict
        ):

            code = (
                normalize_matrix_code(
                    outlet.get(
                        "kode"
                    )
                )
            )

        else:

            code = (
                normalize_matrix_code(
                    outlet
                )
            )

        if code:

            remaining[
                code
            ] = outlet

    route = []

    current = "GUDANG"

    while remaining:

        nearest_code = None

        nearest_distance = float(
            "inf"
        )

        for code, outlet in (
            remaining.items()
        ):

            try:

                distance = (
                    get_matrix_distance(

                        current,

                        code
                    )
                )

            except Exception:

                distance = haversine(

                    ACTIVE[
                        "gudang"
                    ]["lat"],

                    ACTIVE[
                        "gudang"
                    ]["lon"],

                    outlet["lat"],

                    outlet["lon"]
                )

            if (
                distance
                <
                nearest_distance
            ):

                nearest_distance = (
                    distance
                )

                nearest_code = code

        if nearest_code is None:

            break

        route.append(
            nearest_code
        )

        current = nearest_code

        del remaining[
            nearest_code
        ]

    return route


# ============================================================
# LOAD INFORMATION
# ============================================================

def calculate_group_load(
    group
):

    return sum(

        safe_float(
            outlet.get(
                "permintaan"
            )
        )

        for outlet in group
    )


def calculate_utilization(
    load
):

    capacity = safe_float(
        ACTIVE[
            "kapasitas"
        ]
    )

    if capacity <= 0:

        return 0.0

    return (

        load
        /
        capacity
        *
        100
    )


# ============================================================
# SWEEP ORDER
# ============================================================

def calculate_angle(
    outlet,
    warehouse_lat,
    warehouse_lon
):

    lat = safe_float(
        outlet.get(
            "lat"
        )
    )

    lon = safe_float(
        outlet.get(
            "lon"
        )
    )

    return math.atan2(

        lon - warehouse_lon,

        lat - warehouse_lat
    )


def sweep_order(
    outlets
):

    gudang = ACTIVE.get("gudang", {})

    warehouse_lat = to_float(
        gudang.get("lat"),
        None
    )

    warehouse_lon = to_float(
        gudang.get("lon"),
        None
    )

    if warehouse_lat is None or warehouse_lon is None:
        raise ValueError(
            "Koordinat gudang belum tersedia untuk proses optimasi."
        )

    data = []

    for outlet in outlets:

        angle = calculate_angle(

            outlet,

            warehouse_lat,

            warehouse_lon
        )

        distance = haversine(

            warehouse_lat,

            warehouse_lon,

            safe_float(
                outlet.get(
                    "lat"
                )
            ),

            safe_float(
                outlet.get(
                    "lon"
                )
            )
        )

        data.append(

            (
                angle,
                distance,
                outlet
            )
        )

    data.sort(

        key=lambda x: (
            x[0],
            x[1]
        )
    )

    return [

        x[2]

        for x in data
    ]


# ============================================================
# VRP / PEMBAGIAN ARMADA
# ============================================================

def generate_feasible_groups(
    ordered_outlets,
    vehicle_count,
    capacity
):
    """
    VRP:
    Pembagian outlet berdasarkan kapasitas kendaraan.

    Pada tahap ini:
    - Jarak tidak digunakan.
    - Waktu tidak digunakan.
    - Urutan rute belum dioptimasi.
    - Jumlah kendaraan mengikuti input pengguna.
    - Armada tambahan tetap digunakan walaupun utilisasinya rendah.
    """

    if not ordered_outlets:

        return []

    vehicle_count = max(
        1,
        int(vehicle_count)
    )

    capacity = float(
        capacity
    )

    total_demand = sum(

        safe_float(
            outlet.get(
                "permintaan"
            )
        )

        for outlet in ordered_outlets
    )

    if (
        total_demand
        >
        vehicle_count * capacity
        +
        1e-9
    ):

        raise ValueError(

            f"Kapasitas total "
            f"{vehicle_count * capacity:,.0f} kg "
            f"tidak cukup untuk total permintaan "
            f"{total_demand:,.0f} kg. "
            f"Tambahkan armada."
        )

    groups = [
        []
        for _ in range(
            vehicle_count
        )
    ]

    loads = [
        0.0
        for _ in range(
            vehicle_count
        )
    ]

    indexed = list(
        enumerate(
            ordered_outlets
        )
    )

    # Outlet dengan permintaan terbesar
    # ditempatkan terlebih dahulu.
    indexed.sort(

        key=lambda item: (

            -safe_float(
                item[1].get(
                    "permintaan"
                )
            ),

            item[0]
        )
    )

    for _, outlet in indexed:

        demand = safe_float(

            outlet.get(
                "permintaan"
            )
        )

        if (
            demand
            >
            capacity
            +
            1e-9
        ):

            raise ValueError(

                f"Outlet "
                f"{clean_text(outlet.get('kode'))} "
                f"memiliki permintaan "
                f"{demand:,.0f} kg, "
                f"melebihi kapasitas kendaraan "
                f"{capacity:,.0f} kg."
            )

        candidates = [

            i

            for i in range(
                vehicle_count
            )

            if (
                loads[i]
                +
                demand
                <=
                capacity
                +
                1e-9
            )
        ]

        if not candidates:

            raise ValueError(

                f"Outlet "
                f"{clean_text(outlet.get('kode'))} "
                f"tidak dapat ditempatkan "
                f"tanpa melebihi kapasitas kendaraan."
            )

        # Pilih armada dengan muatan paling kecil.
        target = min(

            candidates,

            key=lambda i: (
                loads[i],
                i
            )
        )

        groups[
            target
        ].append(
            outlet
        )

        loads[
            target
        ] += demand

    return groups


def rebalance_groups(
    groups,
    capacity
):

    if not groups:

        return groups

    changed = True

    max_iterations = 100

    iteration = 0

    while (

        changed

        and

        iteration
        <
        max_iterations
    ):

        changed = False

        iteration += 1

        loads = [

            calculate_group_load(
                group
            )

            for group in groups
        ]

        max_index = max(

            range(
                len(groups)
            ),

            key=lambda i:
                loads[i]
        )

        min_index = min(

            range(
                len(groups)
            ),

            key=lambda i:
                loads[i]
        )

        if (

            max_index
            ==
            min_index

            or

            not groups[
                max_index
            ]
        ):

            break

        max_load = loads[
            max_index
        ]

        min_load = loads[
            min_index
        ]

        if max_load <= min_load:

            break

        candidates = []

        for outlet in groups[
            max_index
        ]:

            demand = safe_float(

                outlet.get(
                    "permintaan"
                )
            )

            new_min_load = (

                min_load
                +
                demand
            )

            if (
                new_min_load
                <=
                capacity
            ):

                difference = abs(

                    max_load
                    -
                    new_min_load
                )

                candidates.append(

                    (
                        difference,
                        outlet
                    )
                )

        if not candidates:

            break

        candidates.sort(

            key=lambda x:
                x[0]
        )

        _, selected = (
            candidates[0]
        )

        groups[
            max_index
        ].remove(
            selected
        )

        groups[
            min_index
        ].append(
            selected
        )

        changed = True

        groups = [

            group

            for group in groups

            if group
        ]

    return groups


# ============================================================
# GROUP SCORE
# ============================================================

def evaluate_groups(
    groups
):

    total_distance = 0.0

    total_time = 0.0

    balance_penalty = 0.0

    loads = []

    for group in groups:

        route = nearest_neighbor(
            group
        )

        distance = (
            calculate_route_distance(
                route
            )
        )

        duration = (
            calculate_route_time(
                route
            )
        )

        load = calculate_group_load(
            group
        )

        total_distance += distance

        total_time += duration

        loads.append(
            load
        )

    if loads:

        average_load = (

            sum(loads)
            /
            len(loads)
        )

        balance_penalty = sum(

            abs(

                load
                -
                average_load
            )

            for load in loads
        )

    return (

        total_distance

        +

        (
            balance_penalty
            *
            0.001
        )

        +

        (
            total_time
            *
            0.01
        )
    )


# ============================================================
# PEMBAGIAN ARMADA
# ============================================================

def optimize_vehicle_groups(
    outlets
):
    """
    Tahap VRP:
    Pembagian armada hanya berdasarkan kapasitas.

    NN belum digunakan di sini.
    """

    vehicle_count = max(

        1,

        int(
            ACTIVE[
                "jumlah_kendaraan"
            ]
        )
    )

    capacity = float(

        ACTIVE[
            "kapasitas"
        ]
    )

    return generate_feasible_groups(

        list(outlets),

        vehicle_count,

        capacity
    )


# ============================================================
# HASIL PER ARMADA
# ============================================================

def build_route_data(
    vehicle_number,
    group
):

    # NN menentukan urutan kunjungan.
    route = nearest_neighbor(
        group
    )

    load = calculate_group_load(
        group
    )

    utilization = (
        calculate_utilization(
            load
        )
    )

    distance = (
        calculate_route_distance(
            route
        )
    )

    duration = (
        calculate_route_time(
            route
        )
    )

    # ========================================================
    # BIAYA BBM
    # ========================================================

    # Jarak (km) / km per liter
    bbm_liter = (

        distance
        /
        BBM_KM_PER_LITER
    )

    # Liter x harga BBM
    biaya_bbm = (

        bbm_liter
        *
        ACTIVE[
            "harga_bbm"
        ]
    )

    # ========================================================
    # BIAYA SUPIR
    # ========================================================

    # Tarif supir yang dimasukkan pengguna adalah Rp/jam.
    #
    # Contoh:
    # waktu = 92 menit
    # tarif = Rp25.000/jam
    #
    # biaya = 92 / 60 x 25.000
    #
    biaya_supir_per_jam = safe_float(

        ACTIVE[
            "biaya_supir"
        ]
    )

    biaya_supir = (

        duration
        /
        60.0
        *
        biaya_supir_per_jam
    )

    # ========================================================
    # TOTAL BIAYA
    # ========================================================

    total_biaya = (

        biaya_bbm
        +
        biaya_supir
    )

    return {

        "kendaraan":
            vehicle_number,

        "jumlah_outlet":
            len(group),

        "route":
            route,

        "muatan":
            load,

        "utilisasi":
            utilization,

        "jarak":
            distance,

        "waktu":
            duration,

        "bbm_liter":
            bbm_liter,

        "biaya_bbm":
            biaya_bbm,

        "biaya_supir":
            biaya_supir,

        "total_biaya":
            total_biaya,
    }


# ============================================================
# BUILD RESULT
# ============================================================

def build_result():

    outlets = ACTIVE[
        "outlets"
    ]

    if not outlets:

        raise ValueError(
            "Belum ada data outlet."
        )

    if (

        ACTIVE[
            "matrix_distance"
        ]
        is None

        or

        ACTIVE[
            "matrix_time"
        ]
        is None
    ):

        raise ValueError(

            "Perhitungan jarak dan waktu "
            "belum dilakukan."
        )

    # ========================================================
    # TAHAP 1
    # VRP CAPACITY ONLY
    # ========================================================

    groups = optimize_vehicle_groups(
        outlets
    )

    # ========================================================
    # BASELINE / RUTE AWAL
    #
    # Kelompok armada sudah ditentukan oleh VRP.
    # Namun urutan outlet masih mengikuti urutan hasil
    # pembagian awal.
    #
    # NN BELUM digunakan pada baseline.
    # ========================================================

    initial_routes = []

    for i, group in enumerate(
        groups,
        start=1
    ):

        initial_route = []

        for outlet in group:

            initial_route.append(

                normalize_matrix_code(

                    outlet.get(
                        "kode"
                    )
                )
            )

        load = calculate_group_load(
            group
        )

        utilization = (
            calculate_utilization(
                load
            )
        )

        distance = (
            calculate_route_distance(
                initial_route
            )
        )

        duration = (
            calculate_route_time(
                initial_route
            )
        )

        bbm_liter = (

            distance
            /
            BBM_KM_PER_LITER
        )

        biaya_bbm = (

            bbm_liter
            *
            ACTIVE[
                "harga_bbm"
            ]
        )

        # Biaya supir berdasarkan waktu.
        biaya_supir_per_jam = safe_float(

            ACTIVE[
                "biaya_supir"
            ]
        )

        biaya_supir = (

            duration
            /
            60.0
            *
            biaya_supir_per_jam
        )

        total_biaya = (

            biaya_bbm
            +
            biaya_supir
        )

        initial_routes.append({

            "kendaraan":
                i,

            "jumlah_outlet":
                len(group),

            "route":
                initial_route,

            "muatan":
                load,

            "utilisasi":
                utilization,

            "jarak":
                distance,

            "waktu":
                duration,

            "bbm_liter":
                bbm_liter,

            "biaya_bbm":
                biaya_bbm,

            "biaya_supir":
                biaya_supir,

            "total_biaya":
                total_biaya,
        })

    initial_distance = sum(

        route["jarak"]

        for route
        in initial_routes
    )

    initial_time = sum(

        route["waktu"]

        for route
        in initial_routes
    )

    initial_load = sum(

        route["muatan"]

        for route
        in initial_routes
    )

    initial_bbm = sum(

        route["bbm_liter"]

        for route
        in initial_routes
    )

    initial_biaya_bbm = sum(

        route["biaya_bbm"]

        for route
        in initial_routes
    )

    initial_biaya_supir = sum(

        route["biaya_supir"]

        for route
        in initial_routes
    )

    initial_cost = sum(

        route["total_biaya"]

        for route
        in initial_routes
    )

    # ========================================================
    # TAHAP 2
    # NEAREST NEIGHBOR
    #
    # Kelompok armada TIDAK berubah.
    # Hanya urutan kunjungan yang diubah oleh NN.
    # ========================================================

    optimized_routes = []

    for i, group in enumerate(

        groups,

        start=1
    ):

        route_data = build_route_data(

            i,

            group
        )

        optimized_routes.append(
            route_data
        )

    total_distance = sum(

        route["jarak"]

        for route
        in optimized_routes
    )

    total_time = sum(

        route["waktu"]

        for route
        in optimized_routes
    )

    total_load = sum(

        route["muatan"]

        for route
        in optimized_routes
    )

    total_bbm = sum(

        route["bbm_liter"]

        for route
        in optimized_routes
    )

    total_biaya_bbm = sum(

        route["biaya_bbm"]

        for route
        in optimized_routes
    )

    total_biaya_supir = sum(

        route["biaya_supir"]

        for route
        in optimized_routes
    )

    total_cost = sum(

        route["total_biaya"]

        for route
        in optimized_routes
    )

    average_utilization = 0.0

    if optimized_routes:

        average_utilization = (

            sum(

                route["utilisasi"]

                for route
                in optimized_routes
            )

            /

            len(
                optimized_routes
            )
        )

    # ========================================================
    # SAVING
    # ========================================================

    distance_saving = (

        initial_distance
        -
        total_distance
    )

    distance_saving_percent = 0.0

    if initial_distance > 0:

        distance_saving_percent = (

            distance_saving
            /
            initial_distance
            *
            100.0
        )

    cost_saving = (

        initial_cost
        -
        total_cost
    )

    fuel_saving = (

        initial_bbm
        -
        total_bbm
    )

    result = {

        "initial": {

            "routes":
                initial_routes,

            "total_distance":
                initial_distance,

            "total_time":
                initial_time,

            "total_load":
                initial_load,

            "total_bbm":
                initial_bbm,

            "total_biaya_bbm":
                initial_biaya_bbm,

            "total_biaya_supir":
                initial_biaya_supir,

            "total_cost":
                initial_cost,
        },

        "optimized": {

            "routes":
                optimized_routes,

            "total_distance":
                total_distance,

            "total_time":
                total_time,

            "total_load":
                total_load,

            "total_bbm":
                total_bbm,

            "total_biaya_bbm":
                total_biaya_bbm,

            "total_biaya_supir":
                total_biaya_supir,

            "total_cost":
                total_cost,

            "average_utilization":
                average_utilization,

            "vehicle_count":
                len(
                    optimized_routes
                ),
        },

        "savings": {

            "distance":
                distance_saving,

            "distance_percent":
                distance_saving_percent,

            "fuel_liter":
                fuel_saving,

            "cost":
                cost_saving,
        },

        "meta": {

            "created_at":
                now_text(),

            "capacity":
                ACTIVE[
                    "kapasitas"
                ],

            "harga_bbm":
                ACTIVE[
                    "harga_bbm"
                ],

            "biaya_supir":
                ACTIVE[
                    "biaya_supir"
                ],

            "bbm_km_per_liter":
                BBM_KM_PER_LITER,

            "vrp_rule":
                (
                    "Pembagian armada berdasarkan "
                    "kapasitas kendaraan"
                ),

            "nn_rule":
                (
                    "Penentuan urutan kunjungan "
                    "berdasarkan jarak terdekat"
                ),
        },
    }

    ACTIVE[
        "last_result"
    ] = result

    return result


# ============================================================
# HITUNG MATRIX
# ============================================================

@app.post("/hitung")
def hitung():

    if not ACTIVE[
        "outlets"
    ]:

        flash(
            "Belum ada data outlet.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    try:

        save_history()

        calculate_api_matrix(

            ACTIVE[
                "outlets"
            ]
        )

        flash(

            "Perhitungan jarak dan waktu "
            "berhasil dilakukan.",

            "success"
        )

    except Exception as e:

        flash(

            f"Perhitungan API gagal: {e}",

            "error"
        )

    return redirect(
        url_for("index")
    )


# ============================================================
# OPTIMASI
# ============================================================

@app.post("/optimasi")
def optimasi():

    if not ACTIVE[
        "outlets"
    ]:

        flash(
            "Belum ada data outlet.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    if (

        ACTIVE[
            "matrix_distance"
        ]
        is None

        or

        ACTIVE[
            "matrix_time"
        ]
        is None
    ):

        flash(

            "Silakan lakukan "
            "perhitungan jarak dan waktu "
            "terlebih dahulu.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    try:

        save_history()

        build_result()

        flash(

            "Optimasi rute berhasil dilakukan.",

            "success"
        )

    except Exception as e:

        flash(

            f"Optimasi gagal: {e}",

            "error"
        )

    return redirect(
        url_for("index")
    )


# ============================================================
# TAMBAH OUTLET
# ============================================================

@app.post("/tambah-outlet")
def tambah_outlet():

    kode = clean_text(
        request.form.get(
            "kode"
        )
    )

    nama = clean_text(
        request.form.get(
            "nama"
        )
    )

    alamat = clean_text(
        request.form.get(
            "alamat"
        )
    )

    permintaan = to_float(

        request.form.get(
            "permintaan"
        )
    )

    if not kode:

        flash(

            "Kode outlet wajib diisi.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    if not nama:

        flash(

            "Nama outlet wajib diisi.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    if not alamat:

        flash(

            "Alamat outlet wajib diisi.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    if (

        permintaan is None

        or

        permintaan <= 0
    ):

        flash(

            "Permintaan harus lebih besar dari 0.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    existing_codes = {

        normalize_matrix_code(
            outlet.get(
                "kode"
            )
        )

        for outlet
        in ACTIVE[
            "outlets"
        ]
    }

    if (

        normalize_matrix_code(
            kode
        )

        in

        existing_codes
    ):

        flash(

            f"Kode outlet {kode} sudah digunakan.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    save_history()

    ACTIVE[
        "outlets"
    ].append({

        "kode":
            kode,

        "nama":
            nama,

        "alamat":
            alamat,

        "permintaan":
            float(
                permintaan
            ),

        "lat":
            None,

        "lon":
            None,
    })

    ACTIVE[
        "matrix_distance"
    ] = None

    ACTIVE[
        "matrix_time"
    ] = None

    ACTIVE[
        "matrix_codes"
    ] = []

    ACTIVE[
        "last_api_time"
    ] = None

    ACTIVE[
        "last_result"
    ] = None

    flash(

        f"Outlet {kode} berhasil ditambahkan.",

        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# HAPUS OUTLET
# ============================================================

@app.post("/hapus-outlet/<kode>")
def hapus_outlet(
    kode
):

    kode = clean_text(
        kode
    )

    outlet_found = None

    for outlet in ACTIVE[
        "outlets"
    ]:

        if (

            normalize_matrix_code(
                outlet.get(
                    "kode"
                )
            )

            ==

            normalize_matrix_code(
                kode
            )
        ):

            outlet_found = outlet

            break

    if outlet_found is None:

        flash(

            f"Outlet {kode} tidak ditemukan.",

            "error"
        )

        return redirect(
            url_for("index")
        )

    save_history()

    ACTIVE[
        "outlets"
    ].remove(
        outlet_found
    )

    ACTIVE[
        "matrix_distance"
    ] = None

    ACTIVE[
        "matrix_time"
    ] = None

    ACTIVE[
        "matrix_codes"
    ] = []

    ACTIVE[
        "last_api_time"
    ] = None

    ACTIVE[
        "last_result"
    ] = None

    flash(

        f"Outlet {kode} berhasil dihapus.",

        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# CLEAR DATA
# ============================================================

@app.post("/clear-data")
def clear_data():

    save_history()

    ACTIVE[
        "outlets"
    ] = []

    ACTIVE[
        "matrix_distance"
    ] = None

    ACTIVE[
        "matrix_time"
    ] = None

    ACTIVE[
        "matrix_codes"
    ] = []

    ACTIVE[
        "last_api_time"
    ] = None

    ACTIVE[
        "last_result"
    ] = None

    flash(

        "Seluruh data outlet berhasil dihapus.",

        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# UNDO
# ============================================================

@app.post("/undo")
def undo_action():

    if not HISTORY:

        flash(

            "Tidak ada perubahan "
            "yang dapat dibatalkan.",

            "info"
        )

        return redirect(
            url_for("index")
        )

    current_state = snapshot_state()

    REDO_HISTORY.append(
        current_state
    )

    previous_state = HISTORY.pop()

    restore_state(
        previous_state
    )

    flash(

        "Perubahan terakhir berhasil dibatalkan.",

        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# REDO
# ============================================================

@app.post("/redo")
def redo_action():

    if not REDO_HISTORY:

        flash(

            "Tidak ada perubahan "
            "yang dapat diulangi.",

            "info"
        )

        return redirect(
            url_for("index")
        )

    current_state = snapshot_state()

    HISTORY.append(
        current_state
    )

    next_state = REDO_HISTORY.pop()

    restore_state(
        next_state
    )

    flash(

        "Perubahan berhasil dikembalikan.",

        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# SETTINGS
# ============================================================

@app.post("/settings")
def settings():

    try:

        jumlah_kendaraan = int(

            request.form.get(

                "jumlah_kendaraan",

                ACTIVE[
                    "jumlah_kendaraan"
                ]
            )
        )

        kapasitas = float(

            request.form.get(

                "kapasitas",

                ACTIVE[
                    "kapasitas"
                ]
            )
        )

        harga_bbm = float(

            request.form.get(

                "harga_bbm",

                ACTIVE[
                    "harga_bbm"
                ]
            )
        )

        biaya_supir = float(

            request.form.get(

                "biaya_supir",

                ACTIVE[
                    "biaya_supir"
                ]
            )
        )

        if jumlah_kendaraan <= 0:

            raise ValueError(

                "Jumlah kendaraan "
                "harus lebih besar dari 0."
            )

        if kapasitas <= 0:

            raise ValueError(

                "Kapasitas kendaraan "
                "harus lebih besar dari 0."
            )

        if harga_bbm < 0:

            raise ValueError(

                "Harga BBM tidak boleh negatif."
            )

        if biaya_supir < 0:

            raise ValueError(

                "Biaya supir tidak boleh negatif."
            )

        save_history()

        ACTIVE[
            "jumlah_kendaraan"
        ] = jumlah_kendaraan

        ACTIVE[
            "kapasitas"
        ] = kapasitas

        ACTIVE[
            "harga_bbm"
        ] = harga_bbm

        ACTIVE[
            "biaya_supir"
        ] = biaya_supir

        ACTIVE[
            "last_result"
        ] = None

        flash(

            "Pengaturan berhasil disimpan.",

            "success"
        )

    except Exception as e:

        flash(

            f"Pengaturan gagal disimpan: {e}",

            "error"
        )

    return redirect(
        url_for("index")
    )


# ============================================================
# ROUTE ALIAS UNTUK UI BARU
# ============================================================

@app.post("/simpan-gudang")
def simpan_gudang():

    # Jangan lagi fallback ke gudang Indomaret ketika form kosong.
    # Form wajib mengirim nama dan alamat gudang yang sedang dipilih.
    nama = clean_text(
        request.form.get("nama_gudang")
        or request.form.get("gudang_nama")
        or request.form.get("nama")
    )

    alamat = clean_text(
        request.form.get("alamat_gudang")
        or request.form.get("gudang_alamat")
        or request.form.get("alamat")
    )

    if not nama:
        flash(
            "Nama gudang wajib diisi. Sistem tidak lagi otomatis menggunakan Gudang Distribusi Indomaret.",
            "error"
        )
        return redirect(url_for("index"))

    if not alamat:
        flash(
            "Alamat gudang wajib diisi. Masukkan alamat gudang yang ingin digunakan.",
            "error"
        )
        return redirect(url_for("index"))

    # Dukungan koordinat manual jika UI/form menyediakannya.
    # Ini membuat gudang tetap bisa dipakai walaupun alamat tidak
    # dikenali oleh Nominatim.
    lat = (
        request.form.get("lat_gudang")
        or request.form.get("gudang_lat")
        or request.form.get("latitude_gudang")
        or request.form.get("latitude")
    )

    lon = (
        request.form.get("lon_gudang")
        or request.form.get("gudang_lon")
        or request.form.get("longitude_gudang")
        or request.form.get("longitude")
    )

    try:
        geo = geocode_gudang(
            nama,
            alamat,
            lat,
            lon
        )

        save_history()

        ACTIVE["gudang"] = {
            "nama": nama,
            "alamat": alamat,
            "lat": geo["lat"],
            "lon": geo["lon"],
            "source": geo.get("source", "nominatim"),
            "display_name": geo.get("display_name", ""),
        }

        # Matrix lama wajib dibuang karena titik depot berubah.
        ACTIVE["matrix_distance"] = None
        ACTIVE["matrix_time"] = None
        ACTIVE["matrix_codes"] = []
        ACTIVE["last_api_time"] = None
        ACTIVE["last_result"] = None

        if geo.get("source") == "manual":
            flash(
                "Data gudang berhasil disimpan menggunakan koordinat manual.",
                "success"
            )
        else:
            flash(
                "Data gudang berhasil disimpan dan koordinat berhasil ditemukan otomatis.",
                "success"
            )

    except Exception as e:
        flash(
            f"Data gudang belum disimpan: {e}",
            "error"
        )

    return redirect(
        url_for("index")
    )


@app.post("/simpan-armada")
def simpan_armada():

    try:

        jumlah_kendaraan = int(

            request.form.get(

                "jumlah_kendaraan",

                ACTIVE[
                    "jumlah_kendaraan"
                ]
            )
        )

        kapasitas = float(

            request.form.get(

                "kapasitas",

                ACTIVE[
                    "kapasitas"
                ]
            )
        )

        if jumlah_kendaraan <= 0:

            raise ValueError(

                "Jumlah kendaraan "
                "harus lebih besar dari 0."
            )

        if kapasitas <= 0:

            raise ValueError(

                "Kapasitas kendaraan "
                "harus lebih besar dari 0."
            )

        save_history()

        ACTIVE[
            "jumlah_kendaraan"
        ] = jumlah_kendaraan

        ACTIVE[
            "kapasitas"
        ] = kapasitas

        ACTIVE[
            "last_result"
        ] = None

        flash(

            "Data armada berhasil disimpan.",

            "success"
        )

    except Exception as e:

        flash(

            f"Data armada gagal disimpan: {e}",

            "error"
        )

    return redirect(
        url_for("index")
    )


@app.post("/simpan-pengaturan")
def simpan_pengaturan():

    try:

        harga_bbm = float(

            request.form.get(

                "harga_bbm",

                ACTIVE[
                    "harga_bbm"
                ]
            )
        )

        biaya_supir = float(

            request.form.get(

                "biaya_supir",

                ACTIVE[
                    "biaya_supir"
                ]
            )
        )

        if harga_bbm < 0:

            raise ValueError(

                "Harga BBM tidak boleh negatif."
            )

        if biaya_supir < 0:

            raise ValueError(

                "Biaya supir tidak boleh negatif."
            )

        save_history()

        ACTIVE[
            "harga_bbm"
        ] = harga_bbm

        ACTIVE[
            "biaya_supir"
        ] = biaya_supir

        ACTIVE[
            "last_result"
        ] = None

        flash(

            "Pengaturan biaya berhasil disimpan.",

            "success"
        )

    except Exception as e:

        flash(

            f"Pengaturan gagal disimpan: {e}",

            "error"
        )

    return redirect(
        url_for("index")
    )


@app.post("/simpan-outlet")
def simpan_outlet():

    return tambah_outlet()


@app.post("/hapus-semua-data")
def hapus_semua_data():

    return clear_data()


# ============================================================
# RESET SYSTEM
# ============================================================

@app.post("/reset")
def reset():

    save_history()

    ACTIVE[
        "outlets"
    ] = []

    ACTIVE[
        "matrix_distance"
    ] = None

    ACTIVE[
        "matrix_time"
    ] = None

    ACTIVE[
        "matrix_codes"
    ] = []

    ACTIVE[
        "last_api_time"
    ] = None

    ACTIVE[
        "last_result"
    ] = None

    ACTIVE[
        "jumlah_kendaraan"
    ] = DEFAULT_JUMLAH_KENDARAAN

    ACTIVE[
        "kapasitas"
    ] = DEFAULT_KAPASITAS

    ACTIVE[
        "harga_bbm"
    ] = DEFAULT_HARGA_BBM

    ACTIVE[
        "biaya_supir"
    ] = DEFAULT_BIAYA_SUPIR

    ACTIVE[
        "gudang"
    ] = {

        "nama":
            GUDANG_NAMA,

        "alamat":
            GUDANG_ALAMAT,

        "lat":
            GUDANG_LAT_TETAP,

        "lon":
            GUDANG_LON_TETAP,

        "source":
            "default",

        "display_name":
            GUDANG_ALAMAT,
    }

    flash(

        "Sistem berhasil direset.",

        "success"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# API STATUS
# ============================================================

@app.get("/api/status")
def api_status():

    return {

        "success":
            True,

        "outlet_count":
            len(
                ACTIVE[
                    "outlets"
                ]
            ),

        "matrix_ready":
            (
                ACTIVE[
                    "matrix_distance"
                ]
                is not None
            ),

        "optimization_ready":
            (
                ACTIVE[
                    "last_result"
                ]
                is not None
            ),

        "jumlah_kendaraan":
            ACTIVE[
                "jumlah_kendaraan"
            ],

        "kapasitas":
            ACTIVE[
                "kapasitas"
            ],

        "last_api_time":
            ACTIVE[
                "last_api_time"
            ],
    }


# ============================================================
# API OUTLETS
# ============================================================

@app.get("/api/outlets")
def api_outlets():

    return {

        "success":
            True,

        "count":
            len(
                ACTIVE[
                    "outlets"
                ]
            ),

        "outlets":
            ACTIVE[
                "outlets"
            ],
    }


# ============================================================
# API RESULT
# ============================================================

@app.get("/api/result")
def api_result():

    result = ACTIVE[
        "last_result"
    ]

    if result is None:

        return {

            "success":
                False,

            "message":
                "Belum ada hasil optimasi.",
        }

    return {

        "success":
            True,

        "result":
            result,
    }


# ============================================================
# ERROR HANDLER 413
# ============================================================

@app.errorhandler(413)
def request_entity_too_large(
    error
):

    flash(

        "Ukuran file terlalu besar.",

        "error"
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# ERROR HANDLER 500
# ============================================================

@app.errorhandler(500)
def internal_server_error(
    error
):

    traceback.print_exc()

    return render_template(

        "index.html",

        outlets=ACTIVE[
            "outlets"
        ],

        result=ACTIVE[
            "last_result"
        ],

        gudang=ACTIVE[
            "gudang"
        ],

        import_preview=
            IMPORT_PREVIEW,

        import_total=
            len(
                IMPORT_PREVIEW
            ),

        import_demand=sum(

            safe_float(

                x.get(
                    "permintaan"
                )
            )

            for x
            in IMPORT_PREVIEW
        ),

        jumlah_kendaraan=
            ACTIVE[
                "jumlah_kendaraan"
            ],

        kapasitas=
            ACTIVE[
                "kapasitas"
            ],

        harga_bbm=
            ACTIVE[
                "harga_bbm"
            ],

        biaya_supir=
            ACTIVE[
                "biaya_supir"
            ],

    ), 500


# ============================================================
# EXPORT EXCEL
# ============================================================

def create_export_workbook():

    result = ACTIVE[
        "last_result"
    ]

    if result is None:

        raise ValueError(

            "Belum ada hasil optimasi "
            "yang dapat diekspor."
        )

    output = BytesIO()

    optimized = result[
        "optimized"
    ]

    routes = optimized.get(

        "routes",

        []
    )

    # ========================================================
    # SHEET 1 - REKAP WEB
    # ========================================================

    rekap_rows = []

    for route_data in routes:

        rekap_rows.append({

            "Kendaraan":
                route_data.get(
                    "kendaraan"
                ),

            "Jumlah Outlet":
                route_data.get(
                    "jumlah_outlet"
                ),

            "Rute":
                "Gudang → "
                +
                " → ".join(
                    route_data.get(
                        "route",
                        []
                    )
                )
                +
                " → Gudang",

            "Muatan (kg)":
                route_data.get(
                    "muatan",
                    0
                ),

            "Utilisasi (%)":
                route_data.get(
                    "utilisasi",
                    0
                ),

            "Jarak (km)":
                route_data.get(
                    "jarak",
                    0
                ),

            "Waktu (menit)":
                route_data.get(
                    "waktu",
                    0
                ),

            "BBM (liter)":
                route_data.get(
                    "bbm_liter",
                    0
                ),

            "Biaya BBM":
                route_data.get(
                    "biaya_bbm",
                    0
                ),

            "Biaya Supir":
                route_data.get(
                    "biaya_supir",
                    0
                ),

            "Total Biaya":
                route_data.get(
                    "total_biaya",
                    0
                ),
        })

    rekap_rows.append({

        "Kendaraan":
            "TOTAL",

        "Jumlah Outlet":
            sum(

                route.get(
                    "jumlah_outlet",
                    0
                )

                for route
                in routes
            ),

        "Rute":
            "",

        "Muatan (kg)":
            optimized.get(
                "total_load",
                0
            ),

        "Utilisasi (%)":
            optimized.get(
                "average_utilization",
                0
            ),

        "Jarak (km)":
            optimized.get(
                "total_distance",
                0
            ),

        "Waktu (menit)":
            optimized.get(
                "total_time",
                0
            ),

        "BBM (liter)":
            optimized.get(
                "total_bbm",
                0
            ),

        "Biaya BBM":
            optimized.get(
                "total_biaya_bbm",
                0
            ),

        "Biaya Supir":
            optimized.get(
                "total_biaya_supir",
                0
            ),

        "Total Biaya":
            optimized.get(
                "total_cost",
                0
            ),
    })

    df_rekap = pd.DataFrame(
        rekap_rows
    )

    # ========================================================
    # SHEET 2 - DATA INPUT WEB
    # ========================================================

    input_rows = []

    for outlet in ACTIVE[
        "outlets"
    ]:

        input_rows.append({

            "Kode":
                outlet.get(
                    "kode"
                ),

            "Nama":
                outlet.get(
                    "nama"
                ),

            "Alamat":
                outlet.get(
                    "alamat"
                ),

            "Permintaan (kg)":
                outlet.get(
                    "permintaan"
                ),

            "Latitude":
                outlet.get(
                    "lat"
                ),

            "Longitude":
                outlet.get(
                    "lon"
                ),
        })

    df_input = pd.DataFrame(
        input_rows
    )

    # ========================================================
    # SHEET 3 - HASIL OPTIMASI WEB
    # ========================================================

    hasil_rows = []

    for route_data in routes:

        hasil_rows.append({

            "Kendaraan":
                route_data.get(
                    "kendaraan"
                ),

            "Kode Outlet":
                ", ".join(
                    route_data.get(
                        "route",
                        []
                    )
                ),

            "Jumlah Outlet":
                route_data.get(
                    "jumlah_outlet"
                ),

            "Muatan (kg)":
                route_data.get(
                    "muatan",
                    0
                ),

            "Utilisasi (%)":
                route_data.get(
                    "utilisasi",
                    0
                ),

            "Jarak (km)":
                route_data.get(
                    "jarak",
                    0
                ),

            "Waktu (menit)":
                route_data.get(
                    "waktu",
                    0
                ),

            "BBM (liter)":
                route_data.get(
                    "bbm_liter",
                    0
                ),

            "Biaya BBM":
                route_data.get(
                    "biaya_bbm",
                    0
                ),

            "Biaya Supir":
                route_data.get(
                    "biaya_supir",
                    0
                ),

            "Total Biaya":
                route_data.get(
                    "total_biaya",
                    0
                ),
        })

    df_hasil = pd.DataFrame(
        hasil_rows
    )

    # ========================================================
    # WRITE EXCEL
    # ========================================================

    with pd.ExcelWriter(

        output,

        engine="openpyxl"

    ) as writer:

        df_rekap.to_excel(

            writer,

            sheet_name="REKAP WEB",

            index=False
        )

        df_input.to_excel(

            writer,

            sheet_name="DATA INPUT WEB",

            index=False
        )

        df_hasil.to_excel(

            writer,

            sheet_name="HASIL OPTIMASI WEB",

            index=False
        )

        workbook = writer.book

        for worksheet in (
            workbook.worksheets
        ):

            for column_cells in (
                worksheet.columns
            ):

                max_length = 0

                column_letter = (

                    column_cells[0]
                    .column_letter
                )

                for cell in column_cells:

                    try:

                        cell_length = len(

                            str(
                                cell.value
                            )
                        )

                        if (
                            cell_length
                            >
                            max_length
                        ):

                            max_length = (
                                cell_length
                            )

                    except Exception:

                        pass

                worksheet.column_dimensions[
                    column_letter
                ].width = min(

                    max_length + 2,

                    60
                )

    output.seek(0)

    return output


@app.get("/export-excel")
def export_excel():

    try:

        workbook = (
            create_export_workbook()
        )

        filename = (

            "hasil_optimasi_rute_web_"

            +

            datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )

            +

            ".xlsx"
        )

        return send_file(

            workbook,

            as_attachment=True,

            download_name=filename,

            mimetype=(

                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception as e:

        flash(

            f"Export Excel gagal: {e}",

            "error"
        )

        return redirect(
            url_for("index")
        )


# ============================================================
# DOWNLOAD ALIAS
# ============================================================

@app.get("/download")
def download():

    return export_excel()


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():

    return {

        "status":
            "ok",

        "application":
            "Poltrada Route Optimization System",

        "timestamp":
            now_text(),
    }


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":

    print()

    print(
        "=============================================="
    )

    print(
        "POLTRADA ROUTE OPTIMIZATION SYSTEM"
    )

    print(
        "=============================================="
    )

    print(

        f"Gudang : "
        f"{ACTIVE['gudang'].get('lat')}, "
        f"{ACTIVE['gudang'].get('lon')}"
    )

    print(

        f"Armada : "
        f"{ACTIVE['jumlah_kendaraan']} kendaraan"
    )

    print(

        f"Kapasitas : "
        f"{ACTIVE['kapasitas']} kg/unit"
    )

    print(
        "=============================================="
    )

    # Production/deployment-safe fallback.
    # SnapDeploy dapat menjalankan app.py secara langsung, sehingga
    # aplikasi harus listen pada semua interface dan port dari platform.
    port = int(os.environ.get("PORT", "5000"))
    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
        use_reloader=False
    )
