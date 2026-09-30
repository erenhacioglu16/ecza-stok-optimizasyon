import streamlit as st
import pandas as pd
import numpy as np
from statistics import NormalDist
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Ecza Deposu Stok & Bütçe Paneli", layout="wide")

# --- 1. ŞİFRE EKRANI ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔐 Ecza Deposu Stok Yönetimi - Giriş")
        st.markdown("Devam etmek için lütfen giriş şifrenizi giriniz.")
        
        password = st.text_input("Şifre", type="password")
        if st.button("Sisteme Giriş Yap"):
            if password == "Eren12345":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Hatalı Şifre! Lütfen tekrar deneyiniz.")
        return False
    return True

if check_password():
    st.title("📊 Ecza Deposu Stok Seviyesi & Bütçe Optimizasyon Paneli")
    st.markdown("Excel dosyanızı yükleyin, hedef bütçenizi (ör. 6.2 Milyar TL) girin ve A-B-C-D-E sınıflarına göre emniyet seviyelerini anında hesaplayın.")
    
    # --- 2. DOSYA YÜKLEME ---
    st.markdown("### 📄 1. Excel Dosyasını Yükleyin")
    uploaded_file = st.file_uploader("'Emniyet Seviyesi' Excel dosyasını seçin", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Excel Okuma (Başlıklar 7. satırda (skiprows=6))
            df = pd.read_excel(uploaded_file, sheet_name='Stok seviyesi', skiprows=6)
            
            # Sütun isimlerindeki gizli boşlukları temizleme
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(subset=['Ürün Kodu']).copy()
            
            # Sayısal sütunları temizleme
            numeric_cols = ['Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL', 'Alış Vades']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Mevcut Toplamlar
            mevcut_stok = df['Stok TL'].sum() if 'Stok TL' in df.columns else 0
            mevcut_hedef = df['Hedef TL'].sum() if 'Hedef TL' in df.columns else 0
            mevcut_min = df['Min TL'].sum() if 'Min TL' in df.columns else 0
            mevcut_optimum = df['Optimum TL'].sum() if 'Optimum TL' in df.columns else 0
            mevcut_fazla = df['Fazla TL'].sum() if 'Fazla TL' in df.columns else 0
            
            st.success("✅ Excel dosyası başarıyla yüklendi ve tüm sütunlar okundu.")
            
            # --- 3. MEVCUT DURUM GÖSTERGELERİ ---
            st.markdown("#### 📌 Yüklenen Dosyaya Göre Mevcut Stok Durumu")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Mevcut Stok TL", f"{mevcut_stok:,.0f} ₺".replace(",", "."))
            m2.metric("Mevcut Min TL", f"{mevcut_min:,.0f} ₺".replace(",", "."))
            m3.metric("Mevcut Optimum TL", f"{mevcut_optimum:,.0f} ₺".replace(",", "."))
            m4.metric("Mevcut Hedef TL", f"{mevcut_hedef:,.0f} ₺".replace(",", "."))
            m5.metric("Mevcut Fazla TL", f"{mevcut_fazla:,.0f} ₺".replace(",", "."))
            
            st.markdown("---")
            
            # --- 4. BÜTÇE HEDEFİ VE EMNİYET SEVİYESİ SEÇİMİ ---
            st.markdown("### 🎯 2. Hedef Bütçe & Emniyet Seviyesi Ayarları")
            
            col_b1, col_b2 = st.columns([1, 1])
            with col_b1:
                hedef_butce_input = st.number_input(
                    "Hedeflemek İstediğiniz Toplam Bütçe Sınırı (TL):",
                    value=6200000000,
                    step=50000000,
                    format="%d",
                    help="Örneğin 6.200.000.000 TL (6.2 Milyar) yazabilirsiniz."
                )
            
            st.markdown("**Defa (Pareto) Sınıflarına Göre Emniyet Seviyeleri (Sadece 0 veya 5 ile biten değerler):**")
            
            # A, B, C, D, E Seviye Seçimleri (Sadece 5'in katları: 50, 55, 60, 65, 70, 75, 80, 85, 90, 95)
            options_5 = list(range(50, 100, 5))
            
            p1, p2, p3, p4, p5 = st.columns(5)
            with p1:
                level_A = st.selectbox("A Grubu Emniyet %", options=options_5, index=options_5.index(90))
            with p2:
                level_B = st.selectbox("B Grubu Emniyet %", options=options_5, index=options_5.index(85))
            with p3:
                level_C = st.selectbox("C Grubu Emniyet %", options=options_5, index=options_5.index(80))
            with p4:
                level_D = st.selectbox("D Grubu Emniyet %", options=options_5, index=options_5.index(75))
            with p5:
                level_E = st.selectbox("E Grubu Emniyet %", options=options_5, index=options_5.index(70))
                
            level_map = {'A': level_A, 'B': level_B, 'C': level_C, 'D': level_D, 'E': level_E}
            
            # --- 5. YENİ DEĞERLERİ MATEMATİKSEL OLARAK HESAPLAMA ---
            pareto_col_name = 'Defa Pareto' if 'Defa Pareto' in df.columns else ('Defa ABC' if 'Defa ABC' in df.columns else None)
            
            if pareto_col_name:
                pareto_series = df[pareto_col_name].fillna('E').astype(str).str.strip().str.upper()
            else:
                pareto_series = pd.Series(['E'] * len(df), index=df.index)
                
            df['Yeni_Emniyet_Seviyesi'] = pareto_series.map(level_map).fillna(75)
            
            # NORMSINV Formülü ile Yeni Min, Hedef, Optimum, Fazla TL Hesaplama
            old_levels = df['Emniyet Seviyesi'].clip(50, 99) / 100.0
            new_levels = df['Yeni_Emniyet_Seviyesi'] / 100.0
            
            z_old = old_levels.apply(lambda p: NormalDist().inv_cdf(p))
            z_old = np.where(z_old == 0, 1e-5, z_old)
            z_new = new_levels.apply(lambda p: NormalDist().inv_cdf(p))
            
            # S = Hedef - Min
            df['Hedef_Min_Fark'] = df['Hedef TL'] - df['Min TL']
            
            # Yeni Min TL = Min TL * (Z_new / Z_old)
            df['Yeni Min TL'] = df['Min TL'] * (z_new / z_old)
            # Yeni Hedef TL = Yeni Min TL + Fark
            df['Yeni Hedef TL'] = df['Yeni Min TL'] + df['Hedef_Min_Fark']
            # Yeni Optimum TL = (Yeni Min TL + Yeni Hedef TL) / 2
            df['Yeni Optimum TL'] = (df['Yeni Min TL'] + df['Yeni Hedef TL']) / 2.0
            # Yeni Fazla TL = MAX(0, Stok TL - Yeni Hedef TL)
            df['Yeni Fazla TL'] = np.maximum(0, df['Stok TL'] - df['Yeni Hedef TL'])
            
            # Yeni Toplamlar
            yeni_toplam_hedef = df['Yeni Hedef TL'].sum()
            yeni_toplam_fazla = df['Yeni Fazla TL'].sum()
            
            # --- 6. HESAPLANAN SONUÇ PANOLARI ---
            st.markdown("---")
            st.markdown("### 📈 3. Yeni Belirlenen Hedef Seviyeleri & Bütçe Analizi")
            
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Yeni Hesaplanan Hedef TL", f"{yeni_toplam_hedef:,.0f} ₺".replace(",", "."))
            k2.metric("Girdiğiniz Bütçe Sınırı", f"{hedef_butce_input:,.0f} ₺".replace(",", "."))
            
            fark = hedef_butce_input - yeni_toplam_hedef
            if fark >= 0:
                k3.metric("Kalan Bütçe Payı", f"{fark:,.0f} ₺".replace(",", "."))
                st.success(f"🟢 **BÜTÇE UYGUN:** Belirlediğiniz emniyet seviyeleri ile oluşan Yeni Hedef TL ({yeni_toplam_hedef:,.0f} ₺), {hedef_butce_input:,.0f} ₺ bütçe sınırınızın altında kalmaktadır.")
            else:
                k3.metric("Bütçe Aşım Miktarı", f"{abs(fark):,.0f} ₺".replace(",", "."))
                st.error(f"🔴 **BÜTÇE AŞILDI:** Oluşan Yeni Hedef TL ({yeni_toplam_hedef:,.0f} ₺), girdiğiniz {hedef_butce_input:,.0f} ₺ sınırını {abs(fark):,.0f} ₺ aşıyor. Lütfen emniyet seviyelerini düşürün.")
                
            k4.metric("Yeni Oluşan Fazla (Atıl) Stok TL", f"{yeni_toplam_fazla:,.0f} ₺".replace(",", "."))
            
            # --- 7. TABLO GÖSTERİMİ VE TEMİZ DÜZENLEME ---
            st.markdown("### 📋 4. Ürün Bazlı Detay Tablosu")
            
            gosterim_sutunlari = [
                'Ürün Kodu', 'Ürün', 'Firma', pareto_col_name, 'Emniyet Seviyesi', 
                'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL',
                'Yeni_Emniyet_Seviyesi', 'Yeni Min TL', 'Yeni Optimum TL', 'Yeni Hedef TL', 'Yeni Fazla TL'
            ]
            gosterim_sutunlari = [c for c in gosterim_sutunlari if c in df.columns]
            
            df_display = df[gosterim_sutunlari].copy()
            st.dataframe(df_display, use_container_width=True)
            
            # --- 8. DOĞRUDAN DÜZGÜN EXCEL (.XLSX) OLARAK İNDİRME BUTONU ---
            st.markdown("### 📥 5. Sonuçları Excel (.xlsx) Formatında İndirin")
            st.markdown("Aşağıdaki yeşil butona basarak sonuçları bozulmadan, düzgün Microsoft Excel tablosu olarak bilgisayarınıza indirebilirsiniz.")
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df[gosterim_sutunlari].to_excel(writer, index=False, sheet_name='Emniyet_Optimizasyonu')
                
            st.download_button(
                label="🟢 Sonuçları Excel (.xlsx) Olarak İndir",
                data=excel_buffer.getvalue(),
                file_name=f"Emniyet_Seviyeleri_Optimizasyon_{hedef_butce_input}_TL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Dosya işlenirken bir hata oluştu: {e}")
