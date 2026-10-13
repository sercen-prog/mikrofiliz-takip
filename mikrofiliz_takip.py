import json
import datetime
import io
import openpyxl
import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="🌱 Mikrofiliz Üretim, Satış ve Fire Takip",
    page_icon="🌱",
    layout="wide"
)

# Varsayılan Ürün ve Müşteri Listeleri
DEFAULT_PRODUCTS = [
    "Amaranth Mikrofilizi", "Bezelye Mikrofilizi", "İtalyan Fesleğen Mikrofilizi",
    "Havuç Mikrofilizi", "Kişniş Mikrofilizi", "Mor Reyhan Mikrofilizi",
    "Mor Lahana Mikrofilizi", "Pembe Turp Mikrofilizi", "Mor Turp Mikrofilizi",
    "Roka Mikrofilizi", "Brokoli Mikrofilizi", "Hardal (Mizuna) Mikrofilizi",
    "Beyaz Turp Mikrofilizi", "Thai Fesleğen Mikrofilizi"
]

DEFAULT_CUSTOMERS = [
    "Le Meridyen", "Damlıca Çiftliği", "The Crepe Escape", "Manca", "Mandarin"
]

HEADERS = [
    "Ürün Adı", "Üretim Adeti", "Stoktan Çıkan (Satılan)", "Çöpe Atılan (Fire)", 
    "Kalan Ürün", "Satış Yapılan Yer / Müşteri", "Ekim Günü", "Hasat Süresi (gün)", 
    "Hasat Tarihi", "Karanlıkta (gün)", "Işığa Çıkış Tarihi", "Not"
]

def safe_int(val):
    if not val: return 0
    try:
        val_str = str(val).replace("💡", "").strip()
        return int(float(val_str))
    except:
        return 0

def get_default_sheet_data():
    rows = [HEADERS]
    for p in DEFAULT_PRODUCTS:
        rows.append([p, "0", "0", "0", "0", "", "", "10", "", "3", "", ""])
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
                    sheets[ws.title] = [HEADERS] + clean_prod
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
            tot_u, tot_s, tot_f, tot_k = 0, 0, 0, 0
            for r in data_rows:
                tot_u += safe_int(r[1]) if len(r)>1 else 0
                tot_s += safe_int(r[2]) if len(r)>2 else 0
                tot_f += safe_int(r[3]) if len(r)>3 else 0
                tot_k += safe_int(r[4]) if len(r)>4 else 0
            
            tot_row = ["Toplam Üretim", str(tot_u), str(tot_s), str(tot_f), str(tot_k)] + [""] * (len(HEADERS) - 5)
            full_rows = [HEADERS] + data_rows + [tot_row]
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
                if len(r_copy) > 6: r_copy[6] = format_date_tr(r_copy[6])
                if len(r_copy) > 8: r_copy[8] = format_date_tr(r_copy[8])
                if len(r_copy) > 10: r_copy[10] = format_date_tr(r_copy[10])
                formatted_prod.append(r_copy)
            sheets[s_name] = [HEADERS] + formatted_prod
    return sheets

def export_sheets_to_excel(sheets_data):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    for s_name, rows in sheets_data.items():
        ws = wb.create_sheet(title=s_name)
        data_rows = rows[1:]
        tot_u, tot_s, tot_f, tot_k = 0, 0, 0, 0
        for r in data_rows:
            tot_u += safe_int(r[1]) if len(r)>1 else 0
            tot_s += safe_int(r[2]) if len(r)>2 else 0
            tot_f += safe_int(r[3]) if len(r)>3 else 0
            tot_k += safe_int(r[4]) if len(r)>4 else 0
        
        tot_row = ["Toplam Üretim", str(tot_u), str(tot_s), str(tot_f), str(tot_k)] + [""] * (len(HEADERS) - 5)
        full_rows = [HEADERS] + data_rows + [tot_row]

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

if "custom_customers" not in st.session_state:
    st.session_state.custom_customers = []

parti_names = list(st.session_state.sheets.keys())

if "active_parti" not in st.session_state or st.session_state.active_parti not in parti_names:
    st.session_state.active_parti = parti_names[0]

st.title("🌱 Mikrofiliz Üretim, Satış ve Fire Takip Sistemi")

# 🔔 DÜN / BUGÜN / YARIN BİLDİRİM VE GÖREV PANELİ
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
            u_adet = safe_int(r[1]) if len(r)>1 else 0
            ekim_str = r[6] if len(r)>6 else ""
            hasat_str = r[8] if len(r)>8 else ""
            isik_str = r[10] if len(r)>10 else ""

            if u_adet > 0 and ekim_str:
                isik_dt = parse_tr_date(isik_str)
                hasat_dt = parse_tr_date(hasat_str)

                if hasat_dt == yesterday:
                    dun_islem.append(f"✂️ **{u_name}** ({p_name}) - {u_adet} Adet Hasat Edilmeliydi")
                elif isik_dt == yesterday:
                    dun_islem.append(f"💡 **{u_name}** ({p_name}) - {u_adet} Adet Işığa Çıkarılmalıydı")

                if hasat_dt == today:
                    bugun_islem.append(f"✂️ **{u_name}** ({p_name}) - {u_adet} Adet Hasat Edilecek")
                elif isik_dt == today:
                    bugun_islem.append(f"💡 **{u_name}** ({p_name}) - {u_adet} Adet Işığa Çıkarılacak")

                if hasat_dt == tomorrow:
                    yarin_islem.append(f"✂️ **{u_name}** ({p_name}) - {u_adet} Adet Hasat Edilecek")
                elif isik_dt == tomorrow:
                    yarin_islem.append(f"💡 **{u_name}** ({p_name}) - {u_adet} Adet Işığa Çıkarılacak")

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

# 🛍️ SATIŞ YAP VE MÜŞTERİ KAYIT MODÜLÜ (EN ÜST BÖLÜM)
with st.expander("🛍️ Satış Yap ve Müşteri Yönetimi", expanded=True):
    col_s1, col_s2, col_s3 = st.columns([1, 1, 1.2])
    
    with col_s1:
        s_parti = st.selectbox("📂 Satış Yapılacak Parti", parti_names, key="sales_parti_select")
        
    s_rows = st.session_state.sheets.get(s_parti, [])[1:]
    s_prod_names = [r[0] for r in s_rows if len(r) > 0 and r[0] != "Toplam Üretim"]
    
    with col_s2:
        s_urun = st.selectbox("🌱 Satılacak Ürün", s_prod_names, key="sales_urun_select")
        
    s_idx = next((i for i, r in enumerate(s_rows) if r[0] == s_urun), None)
    
    if s_idx is not None:
        curr_s_row = s_rows[s_idx]
        u_miktari = safe_int(curr_s_row[1]) if len(curr_s_row)>1 else 0
        mevcut_satilan = safe_int(curr_s_row[2]) if len(curr_s_row)>2 else 0
        mevcut_fire = safe_int(curr_s_row[3]) if len(curr_s_row)>3 else 0
        mevcut_kalan = max(0, u_miktari - mevcut_satilan - mevcut_fire)
        ekim_tarihi = format_date_tr(curr_s_row[6]) if len(curr_s_row) > 6 else "-"
        hasat_tarihi = format_date_tr(curr_s_row[8]) if len(curr_s_row) > 8 else "-"
        
        with col_s3:
            st.info(f"📌 **{s_parti} - {s_urun}**\n\nEkim: **{ekim_tarihi}** | Hasat: **{hasat_tarihi}**\n\n**Mevcut Stok:** {mevcut_kalan} Adet (Satılan: {mevcut_satilan} | Fire: {mevcut_fire})")
            
        c_sale1, c_sale2, c_sale3 = st.columns([1, 1.5, 1])
        
        with c_sale1:
            satilacak_adet = st.number_input("Satış Yapılan Adet", min_value=1, max_value=max(1, mevcut_kalan) if mevcut_kalan > 0 else 1, value=1, key="satilacak_adet_input")
            
        all_customers = DEFAULT_CUSTOMERS + st.session_state.custom_customers
        customer_options = all_customers + ["➕ Yeni Müşteri Ekle..."]
        
        with c_sale2:
            selected_customer_option = st.selectbox("Müşteri / Satış Noktası", customer_options, key="customer_select_option")
            
        final_customer = ""
        if selected_customer_option == "➕ Yeni Müşteri Ekle...":
            with c_sale3:
                new_cust_name = st.text_input("Yeni Müşteri Adını Girin", key="new_cust_name_input")
                if st.button("➕ Müşteriyi Kaydet", use_container_width=True):
                    if new_cust_name.strip():
                        if new_cust_name.strip() not in st.session_state.custom_customers and new_cust_name.strip() not in DEFAULT_CUSTOMERS:
                            st.session_state.custom_customers.append(new_cust_name.strip())
                            st.success(f"'{new_cust_name.strip()}' müşteri listesine eklendi!")
                            st.rerun()
            final_customer = new_cust_name.strip() if new_cust_name else ""
        else:
            final_customer = selected_customer_option
            
        if st.button("🛍️ Satışı Onayla ve Tabloya İşle", use_container_width=True):
            if mevcut_kalan < satilacak_adet:
                st.error("Yetersiz stok! Satış adedi kalan stoktan fazla olamaz.")
            elif not final_customer:
                st.error("Lütfen geçerli bir müşteri seçin veya ekleyin.")
            else:
                yeni_satilan_toplam = mevcut_satilan + satilacak_adet
                yeni_kalan = max(0, u_miktari - yeni_satilan_toplam - mevcut_fire)
                
                eski_musteri_metni = curr_s_row[5] if len(curr_s_row) > 5 else ""
                yeni_kayit_metni = f"{final_customer} ({satilacak_adet} Adet)"
                if eski_musteri_metni.strip():
                    yeni_musteri_metni = f"{eski_musteri_metni}, {yeni_kayit_metni}"
                else:
                    yeni_musteri_metni = yeni_kayit_metni
                
                curr_s_row[2] = str(yeni_satilan_toplam)
                curr_s_row[4] = str(yeni_kalan)
                curr_s_row[5] = yeni_musteri_metni
                
                st.session_state.sheets[s_parti][s_idx + 1] = curr_s_row
                save_data_to_sheets(st.session_state.sheets)
                st.success(f"✅ Satış Başarılı! {final_customer} müşterisine {satilacak_adet} adet {s_urun} satıldı.")
                st.rerun()

st.markdown("---")

# 📊 DETAYLI VERİMLİLİK, SATIŞ VE FİRE (ÇÖP) RAPORLAMA MODÜLÜ
with st.expander("📊 Üretim, Satış, Çöp (Fire) ve Müşteri Analiz Raporu", expanded=False):
    report_list = []
    customer_sales = {}
    tot_global_u, tot_global_s, tot_global_f, tot_global_k = 0, 0, 0, 0

    for p_name, p_rows in st.session_state.sheets.items():
        for r in p_rows[1:]:
            if len(r) > 0 and r[0] != "Toplam Üretim":
                u_name = r[0]
                u_adet = safe_int(r[1]) if len(r)>1 else 0
                s_cikan = safe_int(r[2]) if len(r)>2 else 0
                fire = safe_int(r[3]) if len(r)>3 else 0
                kalan = safe_int(r[4]) if len(r)>4 else 0
                musteri = r[5] if len(r)>5 else ""

                tot_global_u += u_adet
                tot_global_s += s_cikan
                tot_global_f += fire
                tot_global_k += kalan

                if u_adet > 0 or s_cikan > 0 or fire > 0 or kalan > 0:
                    report_list.append({
                        "Parti": p_name,
                        "Ürün": u_name,
                        "Üretim": u_adet,
                        "Satılan": s_cikan,
                        "Çöpe Atılan (Fire)": fire,
                        "Kalan": kalan
                    })

                if musteri and s_cikan > 0:
                    for m_item in musteri.split(","):
                        m_item = m_item.strip()
                        if "(" in m_item and ")" in m_item:
                            m_name = m_item.split("(")[0].strip()
                            m_val = safe_int(m_item.split("(")[1].replace("Adet", "").replace(")", "").strip())
                            customer_sales[m_name] = customer_sales.get(m_name, 0) + m_val
                        elif m_item:
                            customer_sales[m_item] = customer_sales.get(m_item, 0) + s_cikan

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🌱 Toplam Üretim", f"{tot_global_u} Adet")
    m2.metric("🛍️ Satılan / Çıkan", f"{tot_global_s} Adet")
    m3.metric("🗑️ Çöpe Atılan (Fire)", f"{tot_global_f} Adet")
    m4.metric("🛒 Depodaki Kalan", f"{tot_global_k} Adet")

    col_g1, col_g2 = st.columns([2, 1])

    with col_g1:
        if report_list:
            df_report = pd.DataFrame(report_list)
            df_grouped = df_report.groupby("Ürün")[["Üretim", "Satılan", "Çöpe Atılan (Fire)", "Kalan"]].sum().reset_index()
            st.markdown("#### 📈 Ürün Bazlı Üretim, Satış, Fire ve Stok Dağılımı")
            st.bar_chart(df_grouped.set_index("Ürün"))
        else:
            st.info("Henüz analiz edilecek veri girilmedi.")

    with col_g2:
        st.markdown("#### 🚚 Müşteri / Satış Noktası Dağılımı")
        if customer_sales:
            df_cust = pd.DataFrame(list(customer_sales.items()), columns=["Müşteri / Yer", "Satılan Adet"]).sort_values(by="Satılan Adet", ascending=False)
            st.dataframe(df_cust, hide_index=True, use_container_width=True)
        else:
            st.info("Henüz müşterilere yapılan satış kaydı girilmedi.")

st.markdown("---")

# YAN MENÜ: Parti Yönetimi, Excel İçe Aktarma & İndirme & Sıfırlama
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
            new_rows = [HEADERS]
            for r in first_sheet[1:]:
                tmpl = [r[0], "0", "0", "0", "0", "", "", r[7] if len(r)>7 else "10", "", r[9] if len(r)>9 else "3", "", ""]
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
    st.header("🧹 Veri Dönüştürme & Temizleme")
    if st.button("🗑️ Mevcut Çıkan Ürünleri Fire'ye Taşı", use_container_width=True):
        for p_key, p_val in st.session_state.sheets.items():
            for r_idx in range(1, len(p_val)):
                r = p_val[r_idx]
                if len(r) > 4:
                    u_val = safe_int(r[1])
                    cikan_val = safe_int(r[2]) # Eski çıkan/satılan
                    mevcut_fire = safe_int(r[3])
                    
                    yeni_fire = mevcut_fire + cikan_val
                    
                    r[2] = "0" # Satılan sıfırlanır
                    r[3] = str(yeni_fire) # Fire'ye eklenir
                    r[4] = str(max(0, u_val - yeni_fire)) # Kalan hesaplanır
                    if len(r) > 5: r[5] = "" # Müşteri temizlenir
                    
        save_data_to_sheets(st.session_state.sheets)
        st.success("Tüm mevcut stoktan çıkan ürünler başarıyla Fire (Çöp) sütununa taşındı!")
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

    if selected_filter_product != "Tüm Tabloyu Göster (Aktif Parti)":
        st.subheader(f"📊 '{selected_filter_product}' İçin Tüm Partilerin Üretim ve Satış Geçmişi")
        
        product_history = []
        tot_u, tot_s, tot_f, tot_k = 0, 0, 0, 0

        for p_name, p_rows in st.session_state.sheets.items():
            for r in p_rows[1:]:
                if len(r) > 0 and r[0] == selected_filter_product:
                    u_adet = safe_int(r[1]) if len(r)>1 else 0
                    s_cikan = safe_int(r[2]) if len(r)>2 else 0
                    fire = safe_int(r[3]) if len(r)>3 else 0
                    kalan = safe_int(r[4]) if len(r)>4 else 0
                    musteri = r[5] if len(r)>5 else ""
                    ekim_str = format_date_tr(r[6]) if len(r)>6 else ""
                    hasat_suresi = r[7] if len(r)>7 else "10"
                    hasat_str = format_date_tr(r[8]) if len(r)>8 else ""
                    karanlik_suresi = r[9] if len(r)>9 else "3"
                    isik_str = format_date_tr(r[10]) if len(r)>10 else ""
                    not_str = r[11] if len(r)>11 else ""

                    tot_u += u_adet
                    tot_s += s_cikan
                    tot_f += fire
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
                        "Üretim Adeti": u_adet,
                        "Satılan": s_cikan,
                        "Çöpe Atılan (Fire)": fire,
                        "Kalan Ürün": kalan,
                        "Satış Yapılan Yer / Müşteri": musteri,
                        "Ekim Günü": ekim_str,
                        "Hasat Tarihi": hasat_str,
                        "Not": not_str
                    })

        if product_history:
            product_history.append({
                "Parti Adı": "📊 GENEL TOPLAM",
                "Durum": "-",
                "Üretim Adeti": tot_u,
                "Satılan": tot_s,
                "Çöpe Atılan (Fire)": tot_f,
                "Kalan Ürün": tot_k,
                "Satış Yapılan Yer / Müşteri": "-",
                "Ekim Günü": "-",
                "Hasat Tarihi": "-",
                "Not": "-"
            })
            st.table(product_history)
        else:
            st.info("Bu ürün için henüz kayıtlı bir üretim verisi bulunamadı.")

    else:
        st.markdown(f"### 📋 {selected_parti} Üretim Tablosu")
        st.markdown("""
        **🎨 Durum Renk Anlamları:** 
        ⚪ **Ekilmedi** | 🩶 **Karanlıkta** | 🟡 **Işığa Çıkmalı 💡** | 🟢 **Hasat Zamanı**
        """)

        table_data = []
        tot_u, tot_s, tot_f, tot_k = 0, 0, 0, 0

        for r in prod_rows:
            u_name = r[0]
            u_adet = safe_int(r[1]) if len(r)>1 else 0
            s_cikan = safe_int(r[2]) if len(r)>2 else 0
            fire = safe_int(r[3]) if len(r)>3 else 0
            kalan = safe_int(r[4]) if len(r)>4 else 0
            musteri = r[5] if len(r)>5 else ""
            ekim_str = format_date_tr(r[6]) if len(r)>6 else ""
            hasat_suresi = r[7] if len(r)>7 else "10"
            hasat_str = format_date_tr(r[8]) if len(r)>8 else ""
            karanlik_suresi = r[9] if len(r)>9 else "3"
            isik_str = format_date_tr(r[10]) if len(r)>10 else ""
            not_str = r[11] if len(r)>11 else ""

            tot_u += u_adet
            tot_s += s_cikan
            tot_f += fire
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
                "Üretim Adeti": u_adet,
                "Satılan": s_cikan,
                "Çöpe Atılan (Fire)": fire,
                "Kalan Ürün": kalan,
                "Satış Yapılan Yer / Müşteri": musteri,
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
            "Üretim Adeti": tot_u,
            "Satılan": tot_s,
            "Çöpe Atılan (Fire)": tot_f,
            "Kalan Ürün": tot_k,
            "Satış Yapılan Yer / Müşteri": "-",
            "Ekim Günü": "-",
            "Karanlıkta (gün)": "-",
            "Işığa Çıkış Tarihi": "-",
            "Hasat Süresi (gün)": "-",
            "Hasat Tarihi": "-",
            "Not": "-"
        })

        st.table(table_data)

    # HIZLI DÜZENLEME PANELİ (SADECE ÜRETİM & FİRE YÖNETİMİ)
    st.markdown("---")
    st.subheader("⚡ Hızlı İşlem Paneli (Ekim & Fire Yönetimi)")

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
                parse_tr_date(curr_r[6]) if len(curr_r)>6 and curr_r[6] else datetime.date.today(), 
                format="DD.MM.YYYY",
                key=f"input_ekim_{target_parti}_{selected_urun}"
            )
            e_u_adet = st.number_input(
                "Üretim Adeti", 
                min_value=0, 
                value=safe_int(curr_r[1]) if len(curr_r)>1 else 0,
                key=f"input_u_adet_{target_parti}_{selected_urun}"
            )
        with c2:
            e_fire = st.number_input(
                "Çöpe Atılan (Fire)", 
                min_value=0, 
                value=safe_int(curr_r[3]) if len(curr_r)>3 else 0,
                key=f"input_fire_{target_parti}_{selected_urun}"
            )
            e_kar = st.number_input(
                "Karanlıkta (gün)", 
                min_value=1, 
                value=safe_int(curr_r[9]) if len(curr_r)>9 and curr_r[9] != "" else 3,
                key=f"input_kar_{target_parti}_{selected_urun}"
            )
        with c3:
            e_has = st.number_input(
                "Hasat Süresi (gün)", 
                min_value=1, 
                value=safe_int(curr_r[7]) if len(curr_r)>7 and curr_r[7] != "" else 10,
                key=f"input_has_{target_parti}_{selected_urun}"
            )
            e_not = st.text_input(
                "Not / Açıklama", 
                value=curr_r[11] if len(curr_r)>11 else "",
                key=f"input_not_{target_parti}_{selected_urun}"
            )

        if st.button(f"⚡ {selected_urun} ({target_parti}) Verilerini Kaydet", use_container_width=True):
            e_s_cikan = safe_int(curr_r[2]) if len(curr_r)>2 else 0
            e_musteri = curr_r[5] if len(curr_r)>5 else ""
            kalan = max(0, e_u_adet - e_s_cikan - e_fire)
            isik_dt = e_ekim + datetime.timedelta(days=int(e_kar))
            hasat_dt = e_ekim + datetime.timedelta(days=int(e_has))

            updated = [
                selected_urun,
                str(e_u_adet),
                str(e_s_cikan),
                str(e_fire),
                str(kalan),
                e_musteri,
                e_ekim.strftime("%d.%m.%Y"),
                str(e_has),
                hasat_dt.strftime("%d.%m.%Y"),
                str(e_kar),
                isik_dt.strftime("%d.%m.%Y"),
                e_not
            ]
            st.session_state.sheets[target_parti][idx + 1] = updated
            save_data_to_sheets(st.session_state.sheets)
            st.success(f"'{selected_urun}' ({target_parti}) güncellendi! Çöpe Atılan: {e_fire} | Kalan Ürün: {kalan}")
            st.rerun()
    else:
        st.warning(f"'{selected_urun}' adlı ürün {target_parti} içerisinde henüz tanımlı değil.")
