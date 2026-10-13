import json
import datetime
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="🌱 Mikrofiliz Üretim ve Stok Takip",
    page_icon="🌱",
    layout="wide"
)

# Google Sheets Bağlantı Kurulumu
@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_json" in st.secrets:
        creds_dict = json.loads(st.secrets["gcp_json"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    elif "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        raise ValueError("Google Service Account anahtarı Streamlit Secrets içinde bulunamadı!")
    return gspread.authorize(creds)

def parse_tr_date(dt_str):
    if not dt_str: return None
    dt_str = str(dt_str).replace("💡", "").strip()
    try:
        if "." in dt_str:
            p = dt_str.split(".")
            return datetime.date(int(p[2]), int(p[1]), int(p[0]))
        elif "-" in dt_str:
            p = dt_str.split("-")
            if len(p[0]) == 4: return datetime.date(int(p[0]), int(p[1]), int(p[2]))
            else: return datetime.date(int(p[2]), int(p[1]), int(p[0]))
    except: return None
    return None

def load_data_from_sheets():
    try:
        gc = get_gspread_client()
        # Google E-Tablo adıyla açıyoruz
        sh = gc.open("Mikrofiliz_Verileri")
        sheets = {}
        for ws in sh.worksheets():
            rows = ws.get_all_values()
            if rows:
                clean_prod = [r for r in rows[1:] if len(r) > 0 and str(r[0]).strip() != "Toplam Üretim"]
                sheets[ws.title] = [rows[0]] + clean_prod
        if sheets:
            return sheets
    except Exception as e:
        st.sidebar.warning(f"Google Sheets yüklenemedi: {e}")

    # Varsayılan şablon
    return {
        "Parti No.1": [
            ["Ürün Adı", "Üretim Adeti", "Stoktan Çıkan", "Kalan Ürün", "Ekim Günü", "Hasat Süresi (gün)", "Hasat Tarihi", "Karanlıkta (gün)", "Işığa Çıkış Tarihi", "Not"],
            ["Amaranth Mikrofilizi", "0", "0", "0", "", "10", "", "3", "", ""],
            ["Bezelye Mikrofilizi", "0", "0", "0", "", "12", "", "4", "", ""],
            ["Roka Mikrofilizi", "0", "0", "0", "", "8", "", "3", "", ""]
        ]
    }

def save_data_to_sheets(sheets_data):
    try:
        gc = get_gspread_client()
        sh = gc.open("Mikrofiliz_Verileri")
        
        existing_worksheets = {ws.title: ws for ws in sh.worksheets()}
        
        for s_name, rows in sheets_data.items():
            if s_name in existing_worksheets:
                ws = existing_worksheets[s_name]
                ws.clear()
            else:
                ws = sh.add_worksheet(title=s_name, rows="100", cols="20")
            
            data_rows = rows[1:]
            tot_u, tot_s, tot_k = 0.0, 0.0, 0.0
            for r in data_rows:
                try: tot_u += float(r[1]) if len(r)>1 and r[1] else 0.0
                except: pass
                try: tot_s += float(r[2]) if len(r)>2 and r[2] else 0.0
                except: pass
                try: tot_k += float(r[3]) if len(r)>3 and r[3] else 0.0
                except: pass
            
            tot_row = ["Toplam Üretim", str(int(tot_u)), str(int(tot_s)), str(int(tot_k))] + [""] * (len(rows[0]) - 4)
            full_rows = [rows[0]] + data_rows + [tot_row]
            ws.update("A1", full_rows)
            
        # Silinen partileri Google Sheets'ten de sil
        for old_title, ws in existing_worksheets.items():
            if old_title not in sheets_data and len(sh.worksheets()) > 1:
                sh.del_worksheet(ws)

        st.toast("✅ Veriler Google Sheets'e kaydedildi!", icon="💾")
    except Exception as e:
        st.error(f"Google Sheets Kayıt Hatası: {e}")

if "sheets" not in st.session_state:
    st.session_state.sheets = load_data_from_sheets()

st.title("🌱 Mikrofiliz Üretim ve Stok Takip Sistemi")

# YAN MENÜ: Parti Yönetimi
with st.sidebar:
    st.header("📂 Parti Yönetimi")
    parti_names = list(st.session_state.sheets.keys())
    selected_parti = st.selectbox("Aktif Parti Seçin", parti_names)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.button("➕ Yeni Parti"):
            nums = [int(p.replace("Parti No.", "").strip()) for p in parti_names if p.startswith("Parti No.")]
            next_num = max(nums) + 1 if nums else 1
            new_name = f"Parti No.{next_num}"
            
            first_sheet = st.session_state.sheets[parti_names[0]]
            new_rows = [first_sheet[0][:]]
            for r in first_sheet[1:]:
                tmpl = [r[0], "0", "0", "0", "", r[5] if len(r)>5 else "10", "", r[7] if len(r)>7 else "3", "", ""]
                new_rows.append(tmpl)
            st.session_state.sheets[new_name] = new_rows
            save_data_to_sheets(st.session_state.sheets)
            st.rerun()

    with col_p2:
        if len(parti_names) > 1:
            if st.button("🗑️ Partiyi Sil"):
                del st.session_state.sheets[selected_parti]
                save_data_to_sheets(st.session_state.sheets)
                st.rerun()

# ANA TABLO
rows = st.session_state.sheets.get(selected_parti, [])
if rows:
    headers = rows[0]
    prod_rows = rows[1:]

    st.markdown("""
    **🎨 Durum Renk Anlamları:** 
    ⚪ **Ekilmedi** | 🩶 **Karanlıkta** | 🟡 **Işığa Çıkmalı 💡** | 🟢 **Hasat Zamanı**
    """)

    today = datetime.date.today()
    table_data = []
    tot_u, tot_s, tot_k = 0.0, 0.0, 0.0

    for r in prod_rows:
        u_name = r[0]
        u_adet = float(r[1]) if len(r)>1 and r[1] else 0.0
        s_cikan = float(r[2]) if len(r)>2 and r[2] else 0.0
        kalan = float(r[3]) if len(r)>3 and r[3] else 0.0
        ekim_str = r[4] if len(r)>4 else ""
        hasat_suresi = r[5] if len(r)>5 else "10"
        hasat_str = r[6] if len(r)>6 else ""
        karanlik_suresi = r[7] if len(r)>7 else "3"
        isik_str = r[8] if len(r)>8 else ""
        not_str = r[9] if len(r)>9 else ""

        tot_u += u_adet
        tot_s += s_cikan
        tot_k += kalan

        isik_dt = parse_tr_date(isik_str)
        hasat_dt = parse_tr_date(hasat_str)

        durum_badge = "⚪ Ekilmedi"
        if hasat_dt and hasat_dt <= today:
            durum_badge = "🟢 HASAT ZAMANI!"
        elif isik_dt and isik_dt <= today:
            durum_badge = "🟡 IŞIĞA ÇIKMALI 💡"
        elif ekim_str:
            durum_badge = "🩶 KARANLIKTA"

        table_data.append({
            "Durum": durum_badge,
            "Ürün Adı": u_name,
            "Üretim Adeti": int(u_adet),
            "Stoktan Çıkan": int(s_cikan),
            "Kalan Ürün": int(kalan),
            "Ekim Günü": ekim_str,
            "Karanlıkta (gün)": karanlik_suresi,
            "Işığa Çıkış Tarihi": isik_str,
            "Hasat Süresi (gün)": hasat_suresi,
            "Hasat Tarihi": hasat_str,
            "Not": not_str
        })

    # Toplam Satırı
    table_data.append({
        "Durum": "📊 TOPLAM",
        "Ürün Adı": "Toplam Üretim",
        "Üretim Adeti": int(tot_u),
        "Stoktan Çıkan": int(tot_s),
        "Kalan Ürün": int(tot_k),
        "Ekim Günü": "-",
        "Karanlıkta (gün)": "-",
        "Işığa Çıkış Tarihi": "-",
        "Hasat Süresi (gün)": "-",
        "Hasat Tarihi": "-",
        "Not": "-"
    })

    st.table(table_data)

    # HIZLI DÜZENLEME PANELİ
    st.markdown("---")
    st.subheader("⚡ Hızlı İşlem Paneli")

    col_u1, col_u2 = st.columns([2, 1])
    with col_u1:
        selected_urun = st.selectbox("Düzenlemek İstediğiniz Ürünü Seçin", [r[0] for r in prod_rows])
    
    idx = next(i for i, r in enumerate(prod_rows) if r[0] == selected_urun)
    curr_r = prod_rows[idx]

    c1, c2, c3 = st.columns(3)
    with c1:
        e_ekim = st.date_input("Ekim Tarihi", parse_tr_date(curr_r[4]) if len(curr_r)>4 and curr_r[4] else datetime.date.today())
        e_u_adet = st.number_input("Üretim Adeti", min_value=0, value=int(float(curr_r[1])) if len(curr_r)>1 and curr_r[1] else 0)
    with c2:
        e_kar = st.number_input("Karanlıkta (gün)", min_value=1, value=int(float(curr_r[7])) if len(curr_r)>7 and curr_r[7] else 3)
        e_s_cikan = st.number_input("Stoktan Çıkan", min_value=0, value=int(float(curr_r[2])) if len(curr_r)>2 and curr_r[2] else 0)
    with c3:
        e_has = st.number_input("Hasat Süresi (gün)", min_value=1, value=int(float(curr_r[5])) if len(curr_r)>5 and curr_r[5] else 10)
        e_not = st.text_input("Not / Açıklama", value=curr_r[9] if len(curr_r)>9 else "")

    if st.button("⚡ Hesapla ve Tabloya İşle", use_container_width=True):
        kalan = max(0, e_u_adet - e_s_cikan)
        isik_dt = e_ekim + datetime.timedelta(days=int(e_kar))
        hasat_dt = e_ekim + datetime.timedelta(days=int(e_has))

        updated = [
            selected_urun,
            str(e_u_adet),
            str(e_s_cikan),
            str(kalan),
            e_ekim.strftime("%d.%m.%Y"),
            str(e_has),
            hasat_dt.strftime("%d.%m.%Y"),
            str(e_kar),
            isik_dt.strftime("%d.%m.%Y"),
            e_not
        ]
        st.session_state.sheets[selected_parti][idx + 1] = updated
        save_data_to_sheets(st.session_state.sheets)
        st.success(f"'{selected_urun}' güncellendi! Işığa Çıkış: {isik_dt.strftime('%d.%m.%Y')} | Hasat: {hasat_dt.strftime('%d.%m.%Y')}")
        st.rerun()
