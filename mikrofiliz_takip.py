import os
import json
import datetime
import streamlit as st

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Damlıca Çiftliği - Mikrofiliz Takip",
    page_icon="🌱",
    layout="wide"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "mikrofiliz_verileri.json")

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"uretimdeki_tepsiler": [], "stok_miktarlari": {}}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

data = load_data()

st.title("🌱 Mikrofiliz Üretim & Stok Takip Paneli")

PRODUCTS = [
    "Amaranth Mikrofilizi", "Bezelye Mikrofilizi", "İtalyan Fesleğen Mikrofilizi",
    "Havuç Mikrofilizi", "Kişniş Mikrofilizi", "Mor Reyhan Mikrofilizi",
    "Mor Lahana Mikrofilizi", "Pembe Turp Mikrofilizi", "Mor Turp Mikrofilizi",
    "Roka Mikrofilizi", "Brokoli Mikrofilizi", "Hardal (Mizuna) Mikrofilizi",
    "Beyaz Turp Mikrofilizi", "Thai Fesleğen Mikrofilizi"
]

tab1, tab2, tab3 = st.tabs(["➕ Yeni Ekim Yap", "📦 Canlı Stok Durumu", "⏳ Seradaki Tepsiler"])

with tab1:
    st.subheader("Yeni Tepsi Ekimi Kaydet")
    col1, col2 = st.columns(2)
    with col1:
        selected_product = st.selectbox("Ürün Çeşidi Seçin", PRODUCTS)
        tepsi_adedi = st.number_input("Ekim Yapılan Tepsi Adedi", min_value=1, max_value=500, value=5)
    with col2:
        ekim_tarihi = st.date_input("Ekim Tarihi", datetime.date.today())
        yetisme_suresi = st.number_input("Tahmini Yetişme Süresi (Gün)", min_value=3, max_value=30, value=10)

    hasat_tarihi = ekim_tarihi + datetime.timedelta(days=int(yetisme_suresi))
    st.info(f"📅 Tahmini Hasat Tarihi: **{hasat_tarihi.strftime('%d.%m.%Y')}**")

    if st.button("🌱 Ekimi Veritabanına Kaydet", use_container_width=True):
        yeni_ekim = {
            "id": len(data["uretimdeki_tepsiler"]) + 1,
            "urun": selected_product,
            "adet": int(tepsi_adedi),
            "ekim_tarihi": ekim_tarihi.strftime("%Y-%m-%d"),
            "hasat_tarihi": hasat_tarihi.strftime("%Y-%m-%d"),
            "asama": "Çimlenmede",
            "durum": "Üretimde"
        }
        data["uretimdeki_tepsiler"].append(yeni_ekim)
        save_data(data)
        st.success(f"{tepsi_adedi} Tepsi {selected_product} ekim kaydı eklendi!")

with tab2:
    st.subheader("📦 Ready-to-Ship / Hasat Edilmiş Stoklar")
    if data["stok_miktarlari"]:
        stok_table = [{"Ürün Adı": k, "Hazır Stok (Tepsi)": f"{v} Adet"} for k, v in data["stok_miktarlari"].items()]
        st.table(stok_table)
    else:
        st.warning("Henüz stokta hazır ürün bulunmuyor.")

    st.markdown("---")
    st.write("🛠️ **Manuel Stok Düzenle**")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        m_urun = st.selectbox("Ürün Seçin", PRODUCTS, key="m_urun")
    with col_s2:
        m_adet = st.number_input("Yeni Stok Adedi", min_value=0, max_value=1000, value=int(data["stok_miktarlari"].get(m_urun, 0)))
    
    if st.button("💾 Stok Adedini Güncelle", use_container_width=True):
        data["stok_miktarlari"][m_urun] = int(m_adet)
        save_data(data)
        st.success(f"{m_urun} stoğu {m_adet} tepsi olarak güncellendi.")
        st.rerun()

with tab3:
    st.subheader("⏳ Serada Yetişen Tepsiler")
    aktif_uretim = [p for p in data["uretimdeki_tepsiler"] if p["durum"] == "Üretimde"]
    if aktif_uretim:
        formatted_list = []
        for p in aktif_uretim:
            formatted_list.append({
                "ID": p["id"],
                "Ürün": p["urun"],
                "Tepsi Adedi": p["adet"],
                "Ekim Tarihi": p["ekim_tarihi"],
                "Tahmini Hasat": p["hasat_tarihi"],
                "Aşama": p["asama"]
            })
        st.table(formatted_list)
        st.markdown("---")
        st.write("✂️ **Hasat İşlemi ve Stoğa Aktarma**")
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            hasat_id = st.selectbox("Hasat Edilecek Partinin ID Numarası", [p["id"] for p in aktif_uretim])
        if st.button("✅ Partiyi Hasat Et ve Stoğa Ekle", use_container_width=True):
            for p in data["uretimdeki_tepsiler"]:
                if p["id"] == hasat_id:
                    p["durum"] = "Hasat Edildi"
                    mevcut_stok = data["stok_miktarlari"].get(p["urun"], 0)
                    data["stok_miktarlari"][p["urun"]] = mevcut_stok + p["adet"]
                    break
            save_data(data)
            st.success("Tepsiler başarıyla hasat edildi ve canlı stoğa eklendi!")
            st.rerun()
    else:
        st.info("Serada şu an aktif üretimde olan tepsi bulunmuyor.")