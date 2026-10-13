import os
import datetime
import openpyxl
import streamlit as st

st.set_page_config(
    page_title="🌱 Mikrofiliz Üretim ve Stok Takip",
    page_icon="🌱",
    layout="wide"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "Güncel Üretim.xlsx")

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

def format_date_tr(dt_val):
    if not dt_val: return ""
    if isinstance(dt_val, (datetime.date, datetime.datetime)):
        return dt_val.strftime("%d.%m.%Y")
    return str(dt_val)

def load_excel_data():
    sheets = {}
    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
            for s_name in wb.sheetnames:
                ws = wb[s_name]
                rows_data = []
                for row in ws.iter_rows(values_only=True):
                    row_vals = [str(cell) if cell is not None else "" for cell in row]
                    if any(row_vals):
                        rows_data.append(row_vals)
                if rows_data:
                    clean_prod = [r for r in rows_data[1:] if len(r) > 0 and str(r[0]).strip() != "Toplam Üretim"]
                    sheets[s_name] = [rows_data[0]] + clean_prod
        except Exception as e:
            st.error(f"Excel Okuma Hatası: {e}")
    
    if not sheets:
        sheets["Parti No.1"] = [
            ["Ürün Adı", "Üretim Adeti", "Stoktan Çıkan", "Kalan Ürün", "Ekim Günü", "Hasat Süresi (gün)", "Hasat Tarihi", "Karanlıkta (gün)", "Işığa Çıkış Tarihi", "Not"],
            ["Amaranth Mikrofilizi", "0", "0", "0", "", "10", "", "3", "", ""],
            ["Bezelye Mikrofilizi", "0", "0", "0", "", "12", "", "4", "", ""],
            ["Roka Mikrofilizi", "0", "0", "0", "", "8", "", "3", "", ""]
        ]
    return sheets

def save_excel_data(sheets_data):
    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for s_name, rows in sheets_data.items():
            ws = wb.create_sheet(title=s_name)
            # Toplam hesaplayarak yaz
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

            for r_idx, r_vals in enumerate(full_rows, 1):
                for c_idx, val in enumerate(r_vals, 1):
                    val_str = str(val).replace("💡", "").strip()
                    try:
                        num_v = float(val_str)
                        ws.cell(row=r_idx, column=c_idx, value=int(num_v) if num_v.is_integer() else num_v)
                    except:
                        ws.cell(row=r_idx, column=c_idx, value=val_str)
        wb.save(EXCEL_FILE)
    except Exception as e:
        st.error(f"Excel Kayıt Hatası: {e}")

if "sheets" not in st.session_state:
    st.session_state.sheets = load_excel_data()

st.title("🌱 Mikrofiliz Üretim ve Stok Takip Sistemi")

# SOL MENÜ: Parti Yönetimi
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
                # Ürün adı ve karanlık/hasat sürelerini koru, miktarları ve tarihleri sıfırla
                tmpl = [r[0], "0", "0", "0", "", r[5] if len(r)>5 else "10", "", r[7] if len(r)>7 else "3", "", ""]
                new_rows.append(tmpl)
            st.session_state.sheets[new_name] = new_rows
            save_excel_data(st.session_state.sheets)
            st.rerun()

    with col_p2:
        if len(parti_names) > 1:
            if st.button("🗑️ Partiyi Sil"):
                del st.session_state.sheets[selected_parti]
                save_excel_data(st.session_state.sheets)
                st.rerun()

# ANA TABLO VE RENKLENDİRME
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

    # Toplam Satırını Tabloya Ekle
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
        save_excel_data(st.session_state.sheets)
        st.success(f"'{selected_urun}' güncellendi! Işığa Çıkış: {isik_dt.strftime('%d.%m.%Y')} | Hasat: {hasat_dt.strftime('%d.%m.%Y')}")
        st.rerun()
