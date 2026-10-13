import json
import datetime
import io
import openpyxl
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="🌱 Mikrofiliz Üretim ve Stok Takip",
    page_icon="🌱",
    layout="wide"
)

# Varsayılan Ürün Listesi ve Şablon
DEFAULT_PRODUCTS = [
    "Amaranth Mikrofilizi", "Bezelye Mikrofilizi", "İtalyan Fesleğen Mikrofilizi",
    "Havuç Mikrofilizi", "Kişniş Mikrofilizi", "Mor Reyhan Mikrofilizi",
    "Mor Lahana Mikrofilizi", "Pembe Turp Mikrofilizi", "Mor Turp Mikrofilizi",
    "Roka Mikrofilizi", "Brokoli Mikrofilizi", "Hardal (Mizuna) Mikrofilizi",
    "Beyaz Turp Mikrofilizi", "Thai Fesleğen Mikrofilizi"
]

def get_default_sheet_data():
    headers = ["Ürün Adı", "Üretim Adeti", "Stoktan Çıkan", "Kalan Ürün", "Ekim Günü", "Hasat Süresi (gün)", "Hasat Tarihi", "Karanlıkta (gün)", "Işığa Çıkış Tarihi", "Not"]
    rows = [headers]
    for p in DEFAULT_PRODUCTS:
        rows.append([p, "0", "0", "0", "", "10", "", "3", "", ""])
    return {"Parti No.1": rows}

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

def format_date_tr(dt_val):
    if not dt_val: return ""
    if isinstance(dt_val, (datetime.date, datetime.datetime)):
        return dt_val.strftime("%d.%m.%Y")
    parsed = parse_tr_date(dt_val)
    if parsed:
        return parsed.strftime("%d.%m.%Y")
    return str(dt_val)

def load_data_from_sheets():
    try:
        gc = get_gspread_client()
        sh = gc.open("Mikrofiliz_Verileri")
        sheets = {}
        for ws in sh.worksheets():
            rows = ws.get_all_values()
            if len(rows) > 1:
                clean_prod = [r for r in rows[1:] if len(r) > 0 and str(r[0]).strip() != "Toplam Üretim"]
                if clean_prod:
                    sheets[ws.title] = [rows[0]] + clean_prod
        if sheets:
            return sheets
    except Exception as e:
        st.sidebar.warning(f"Google Sheets yüklenemedi: {e}")

    return get_default_sheet_data()

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
            
        for old_title, ws in existing_worksheets.items():
            if old_title not in sheets_data and len(sh.worksheets()) > 1:
                sh.del_worksheet(ws)

        st.toast("✅ Veriler Google Sheets'e kaydedildi!", icon="💾")
    except Exception as e:
        st.error(f"Google Sheets Kayıt Hatası: {e}")

def parse_excel_file(uploaded_file):
    sheets = {}
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    for s_name in wb.sheetnames:
        ws = wb[s_name]
        rows_data = []
        for row in ws.iter_rows(values_only=True):
            row_vals = [str(cell) if cell is not None else "" for cell in row]
            if any(row_vals):
                rows_data.append(row_vals)
        if len(rows_data) > 1:
            clean_prod = [r for r in rows_data[1:] if len(r) > 0 and str(r[0]).strip() != "Toplam Üretim"]
            formatted_prod = []
            for r in clean_prod:
                r_copy = list(r)
                if len(r_copy) > 4: r_copy[4] = format_date_tr(r_copy[4])
                if len(r_copy) > 6: r_copy[6] = format_date_tr(r_copy[6])
                if len(r_copy) > 8: r_copy[8] = format_date_tr(r_copy[8])
                formatted_prod.append(r_copy)
            sheets[s_name] = [rows_data[0]] + formatted_prod
    return sheets

def export_sheets_to_excel(sheets_data):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    for s_name, rows in sheets_data.items():
        ws = wb.create_sheet(title=s_name)
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

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

if "sheets" not in st.session_state:
    st.session_state.sheets = load_data_from_sheets()

parti_names = list(st.session_state.sheets.keys())

if "active_parti" not in st.session_state or st.session_state.active_parti not in parti_names:
    st.session_state.active_parti = parti_names[0]

st.title("🌱 Mikrofiliz Üretim ve Stok Takip Sistemi")

# 🔔 DÜN / BUGÜN / YARIN BİLDİRİM VE GÖREV PANELİ (TARİHLİ)
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
tomorrow = today + datetime.timedelta(days=1)

str_yesterday = yesterday.strftime("%d.%m.%Y")
str_today = today.strftime("%d.%m.%Y")
str_tomorrow = tomorrow.strftime("%d.%m.%Y")

dun_islem = []
bugun_islem = []
yarin_islem = []

for p_name, p_rows in st.session_state.sheets.items():
    for r in p_rows[1:]:
        if len(r) > 0:
            u_name = r[0]
            u_adet = float(r[1]) if len(r)>1 and r[1] else 0.0
            ekim_str = r[4] if len(r)>4 else ""
            hasat_str = r[6] if len(r)>6 else ""
            isik_str = r[8] if len(r)>8 else ""

            if u_adet > 0 and ekim_str:
                isik_dt = parse_tr_date(isik_str)
                hasat_dt = parse_tr_date(hasat_str)

                # DÜN
                if hasat_dt == yesterday:
                    dun_islem.append(f"✂️ **{u_name}** ({p_name}) - {int(u_adet)} Adet Hasat Edilmeliydi")
                elif isik_dt == yesterday:
                    dun_islem.append(f"💡 **{u_name}** ({p_name}) - {int(u_adet)} Adet Işığa Çıkarılmalıydı")

                # BUGÜN
                if hasat_dt == today:
                    bugun_islem.append(f"✂️ **{u_name}** ({p_name}) - {int(u_adet)} Adet Hasat Edilecek")
                elif isik_dt == today:
                    bugun_islem.append(f"💡 **{u_name}** ({p_name}) - {int(u_adet)} Adet Işığa Çıkarılacak")

                # YARIN
                if hasat_dt == tomorrow:
                    yarin_islem.append(f"✂️ **{u_name}** ({p_name}) - {int(u_adet)} Adet Hasat Edilecek")
                elif isik_dt == tomorrow:
                    yarin_islem.append(f"💡 **{u_name}** ({p_name}) - {int(u_adet)} Adet Işığa Çıkarılacak")

if dun_islem or bugun_islem or yarin_islem:
    st.markdown("### 🔔 Görev & Takip Paneli")
    col_d, col_b, col_y = st.columns(3)
    
    with col_d:
        if dun_islem:
            st.error(f"🔴 **DÜN ({str_yesterday}):**\n\n" + "\n\n".join(dun_islem))
        else:
            st.info(f"🔴 **DÜN ({str_yesterday}):**\nGeciken işlem yok.")

    with col_b:
        if bugun_islem:
            st.warning(f"🟡 **BUGÜN ({str_today}):**\n\n" + "\n\n".join(bugun_islem))
        else:
            st.success(f"🟡 **BUGÜN ({str_today}):**\nAcil işlem yok.")

    with col_y:
        if yarin_islem:
            st.info(f"🔵 **YARIN ({str_tomorrow}):**\n\n" + "\n\n".join(yarin_islem))
        else:
            st.info(f"🔵 **YARIN ({str_tomorrow}):**\nPlanlı işlem yok.")
            
    st.markdown("---")

# YAN MENÜ: Parti Yönetimi, Excel İçe Aktarma & İndirme
with st.sidebar:
    st.header("📂 Parti Yönetimi")
    
    selected_sidebar_parti = st.selectbox(
        "Aktif Parti Seçin", 
        parti_names, 
        index=parti_names.index(st.session_state.active_parti),
        key="sb_parti_select"
    )
    if selected_sidebar_parti != st.session_state.active_parti:
        st.session_state.active_parti = selected_sidebar_parti
        st.rerun()

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
            st.session_state.active_parti = new_name
            save_data_to_sheets(st.session_state.sheets)
            st.rerun()

    with col_p2:
        if len(parti_names) > 1:
            if st.button("🗑️ Partiyi Sil"):
                del st.session_state.sheets[st.session_state.active_parti]
                st.session_state.active_parti = list(st.session_state.sheets.keys())[0]
                save_data_to_sheets(st.session_state.sheets)
                st.rerun()

    st.markdown("---")
    st.header("📥 Excel Olarak İndir")
    excel_buffer = export_sheets_to_excel(st.session_state.sheets)
    st.download_button(
        label="📊 Güncel Verileri Excel Olarak İndir",
        data=excel_buffer,
        file_name=f"Mikrofiliz_Uretim_Takip_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.markdown("---")
    st.header("📤 Excel İçe Aktar")
    uploaded_excel = st.file_uploader("Güncel Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx"])
    if uploaded_excel is not None:
        if st.button("🔄 Excel Verilerini Aktar ve Kaydet", use_container_width=True):
            try:
                parsed_sheets = parse_excel_file(uploaded_excel)
                if parsed_sheets:
                    st.session_state.sheets = parsed_sheets
                    st.session_state.active_parti = list(parsed_sheets.keys())[0]
                    save_data_to_sheets(parsed_sheets)
                    st.success("Excel başarıyla okundu ve Google Sheets veritabanına aktarıldı!")
                    st.rerun()
                else:
                    st.error("Excel dosyasından geçerli sayfa verisi okunamadı.")
            except Exception as e:
                st.error(f"Excel Okuma Hatası: {e}")

# ÜRÜN BAZLI FİLTRELEME VE GÖRÜNTÜLEME MODU
selected_parti = st.session_state.active_parti
rows = st.session_state.sheets.get(selected_parti, [])

if rows:
    prod_rows = rows[1:]
    all_product_names = sorted(list(set([r[0] for sheet in st.session_state.sheets.values() for r in sheet[1:] if len(r)>0])))

    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        selected_filter_product = st.selectbox(
            "🔍 Ürün Bazlı İnceleme (Tüm Partiler Genelinde):",
            ["Tüm Tabloyu Göster (Aktif Parti)"] + all_product_names
        )

    # EĞER BELİRLİ BİR ÜRÜN SEÇİLDİYSE
    if selected_filter_product != "Tüm Tabloyu Göster (Aktif Parti)":
        st.subheader(f"📊 '{selected_filter_product}' İçin Tüm Partilerin Üretim Geçmişi")
        
        product_history = []
        tot_u, tot_s, tot_k = 0.0, 0.0, 0.0

        for p_name, p_rows in st.session_state.sheets.items():
            for r in p_rows[1:]:
                if len(r) > 0 and r[0] == selected_filter_product:
                    u_adet = float(r[1]) if len(r)>1 and r[1] else 0.0
                    s_cikan = float(r[2]) if len(r)>2 and r[2] else 0.0
                    kalan = float(r[3]) if len(r)>3 and r[3] else 0.0
                    ekim_str = format_date_tr(r[4]) if len(r)>4 else ""
                    hasat_suresi = r[5] if len(r)>5 else "10"
                    hasat_str = format_date_tr(r[6]) if len(r)>6 else ""
                    karanlik_suresi = r[7] if len(r)>7 else "3"
                    isik_str = format_date_tr(r[8]) if len(r)>8 else ""
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

                    product_history.append({
                        "Parti Adı": p_name,
                        "Durum": durum_badge,
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

        if product_history:
            product_history.append({
                "Parti Adı": "📊 GENEL TOPLAM",
                "Durum": "-",
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
            st.table(product_history)
        else:
            st.info("Bu ürün için henüz kayıtlı bir üretim verisi bulunamadı.")

    # NORMAL TÜM PARTİ TABLOSU GÖRÜNÜMÜ
    else:
        st.markdown(f"### 📋 {selected_parti} Üretim Tablosu")
        st.markdown("""
        **🎨 Durum Renk Anlamları:** 
        ⚪ **Ekilmedi** | 🩶 **Karanlıkta** | 🟡 **Işığa Çıkmalı 💡** | 🟢 **Hasat Zamanı**
        """)

        table_data = []
        tot_u, tot_s, tot_k = 0.0, 0.0, 0.0

        for r in prod_rows:
            u_name = r[0]
            u_adet = float(r[1]) if len(r)>1 and r[1] else 0.0
            s_cikan = float(r[2]) if len(r)>2 and r[2] else 0.0
            kalan = float(r[3]) if len(r)>3 and r[3] else 0.0
            ekim_str = format_date_tr(r[4]) if len(r)>4 else ""
            hasat_suresi = r[5] if len(r)>5 else "10"
            hasat_str = format_date_tr(r[6]) if len(r)>6 else ""
            karanlik_suresi = r[7] if len(r)>7 else "3"
            isik_str = format_date_tr(r[8]) if len(r)>8 else ""
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

    col_hp1, col_hp2 = st.columns(2)
    
    with col_hp1:
        selected_urun = st.selectbox("🌱 İşlem Yapılacak Ürün", all_product_names, key="quick_urun_select")

    with col_hp2:
        target_parti = st.selectbox(
            "📂 İşlem Yapılacak Parti",
            parti_names,
            index=parti_names.index(selected_parti) if selected_parti in parti_names else 0,
            key="quick_parti_select"
        )

    target_sheet_rows = st.session_state.sheets.get(target_parti, [])[1:]
    idx = next((i for i, r in enumerate(target_sheet_rows) if r[0] == selected_urun), None)

    if idx is not None:
        curr_r = target_sheet_rows[idx]

        c1, c2, c3 = st.columns(3)
        with c1:
            e_ekim = st.date_input(
                "Ekim Tarihi (Gün/Ay/Yıl)", 
                parse_tr_date(curr_r[4]) if len(curr_r)>4 and curr_r[4] else datetime.date.today(), 
                format="DD.MM.YYYY",
                key=f"input_ekim_{target_parti}_{selected_urun}"
            )
            e_u_adet = st.number_input(
                "Üretim Adeti", 
                min_value=0, 
                value=int(float(curr_r[1])) if len(curr_r)>1 and curr_r[1] else 0,
                key=f"input_u_adet_{target_parti}_{selected_urun}"
            )
        with c2:
            e_kar = st.number_input(
                "Karanlıkta (gün)", 
                min_value=1, 
                value=int(float(curr_r[7])) if len(curr_r)>7 and curr_r[7] else 3,
                key=f"input_kar_{target_parti}_{selected_urun}"
            )
            e_s_cikan = st.number_input(
                "Stoktan Çıkan", 
                min_value=0, 
                value=int(float(curr_r[2])) if len(curr_r)>2 and curr_r[2] else 0,
                key=f"input_s_cikan_{target_parti}_{selected_urun}"
            )
        with c3:
            e_has = st.number_input(
                "Hasat Süresi (gün)", 
                min_value=1, 
                value=int(float(curr_r[5])) if len(curr_r)>5 and curr_r[5] else 10,
                key=f"input_has_{target_parti}_{selected_urun}"
            )
            e_not = st.text_input(
                "Not / Açıklama", 
                value=curr_r[9] if len(curr_r)>9 else "",
                key=f"input_not_{target_parti}_{selected_urun}"
            )

        if st.button(f"⚡ {selected_urun} ({target_parti}) Verilerini Kaydet", use_container_width=True):
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
            st.session_state.sheets[target_parti][idx + 1] = updated
            save_data_to_sheets(st.session_state.sheets)
            st.success(f"'{selected_urun}' ({target_parti}) güncellendi! Işığa Çıkış: {isik_dt.strftime('%d.%m.%Y')} | Hasat: {hasat_dt.strftime('%d.%m.%Y')}")
            st.rerun()
    else:
        st.warning(f"'{selected_urun}' adlı ürün {target_parti} içerisinde henüz tanımlı değil.")
