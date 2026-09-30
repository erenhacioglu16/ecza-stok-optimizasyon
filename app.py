import streamlit as st
import pandas as pd
import numpy as np
from statistics import NormalDist
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Ecza Deposu Stok & Bütçe Paneli", layout="wide")

# Session State Varsayılan Değerlerini İlklendirme
default_states = {
    "level_A": 90,
    "level_B": 85,
    "level_C": 80,
    "level_D": 75,
    "level_E": 70,
    "level_BOS": 50,
    "hedef_butce": 6200000000,
    "form_submitted": False
}
for key, val in default_states.items():
    if key not in st.session_state:
        st.session_state[key] = val

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
    st.markdown("Excel dosyanızı yükleyin, hedef bütçenizi girin ve Pareto sınıflarına göre emniyet seviyelerini anında hesaplayın.")
    
    # --- 2. DOSYA YÜKLEME ---
    st.markdown("### 📄 1. Excel Dosyasını Yükleyin")
    uploaded_files = st.file_uploader(
        "'Emniyet Seviyesi' Excel dosyasını seçin (Lütfen tek dosya seçiniz)", 
        type=["xlsx", "xls", "xlsb"],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        uploaded_file = uploaded_files[-1]
        try:
            engine_choice = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else None
            df = pd.read_excel(uploaded_file, sheet_name='Stok seviyesi', skiprows=6, engine=engine_choice)
            
            # Sütun isimlerini ve verileri temizleme
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(subset=['Ürün Kodu']).copy()
            
            numeric_cols = ['Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL', 'Alış Vades']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Pareto Sütun tespiti ve Boş/Tanımsız Temizliği
            pareto_col_name = 'Defa Pareto' if 'Defa Pareto' in df.columns else ('Defa ABC' if 'Defa ABC' in df.columns else None)
            if pareto_col_name:
                pareto_series = df[pareto_col_name].fillna('BOŞ').astype(str).str.strip().str.upper()
                pareto_series = pareto_series.replace(['NAN', 'NONE', ''], 'BOŞ')
            else:
                pareto_series = pd.Series(['BOŞ'] * len(df), index=df.index)
            df['Temiz_Pareto'] = pareto_series

            # Mevcut Toplamlar
            mevcut_stok = df['Stok TL'].sum() if 'Stok TL' in df.columns else 0
            mevcut_hedef = df['Hedef TL'].sum() if 'Hedef TL' in df.columns else 0
            mevcut_min = df['Min TL'].sum() if 'Min TL' in df.columns else 0
            mevcut_optimum = df['Optimum TL'].sum() if 'Optimum TL' in df.columns else 0
            mevcut_fazla = df['Fazla TL'].sum() if 'Fazla TL' in df.columns else 0
            
            st.success(f"✅ **{uploaded_file.name}** dosyası başarıyla yüklendi ve işlendi.")
            
            # --- MEVCUT DURUM METRİKLERİ ---
            st.markdown("#### 📌 Yüklenen Dosyaya Göre Mevcut Stok Durumu")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Mevcut Stok TL", f"{mevcut_stok:,.0f} ₺".replace(",", "."))
            m2.metric("Mevcut Min TL", f"{mevcut_min:,.0f} ₺".replace(",", "."))
            m3.metric("Mevcut Optimum TL", f"{mevcut_optimum:,.0f} ₺".replace(",", "."))
            m4.metric("Mevcut Hedef TL", f"{mevcut_hedef:,.0f} ₺".replace(",", "."))
            m5.metric("Mevcut Fazla TL", f"{mevcut_fazla:,.0f} ₺".replace(",", "."))
            
            st.markdown("---")

            # --- 3. AKILLI BÜTÇE & PARETO ÖNERİ ASİSTANI ---
            with st.expander("💡 🤖 Bütçeye Göre Akıllı Pareto Öneri Asistanı (Otomatik Senaryo Analizi)", expanded=False):
                st.markdown("Hedeflemek istediğiniz bütçe sınırını girin. Sistem, Pareto sınıflarına uygun en mantıklı hizmet seviyesi kombinasyonlarını hesaplayıp size alternatifler sunacaktır.")
                
                col_rec1, col_rec2 = st.columns([2, 1])
                with col_rec1:
                    analiz_butce = st.number_input(
                        "Analiz Edilecek Hedef Bütçe Sınırı (TL):",
                        value=int(st.session_state["hedef_butce"]),
                        step=50000000,
                        format="%d",
                        key="analiz_butce_input"
                    )
                with col_rec2:
                    st.write("")
                    st.write("")
                    run_analysis = st.button("🔍 Senaryoları Analiz Et ve Getir", use_container_width=True)

                if run_analysis:
                    # Hızlı Z-Score ve Grup Toplamları Hesabı
                    old_levels_all = df['Emniyet Seviyesi'].clip(50, 99) / 100.0
                    z_old_all = old_levels_all.apply(lambda p: NormalDist().inv_cdf(p))
                    z_old_all = np.where(z_old_all == 0, 1e-5, z_old_all)
                    
                    df_temp = pd.DataFrame({
                        'cls': df['Temiz_Pareto'],
                        'min_div_z': df['Min TL'] / z_old_all,
                        'fark': df['Hedef TL'] - df['Min TL']
                    })
                    
                    grouped_stats = df_temp.groupby('cls').agg(
                        sum_min_z=('min_div_z', 'sum'),
                        sum_fark=('fark', 'sum')
                    ).to_dict(orient='index')

                    options_5 = list(range(50, 100, 5))
                    combos = []
                    for A in options_5:
                        for B in [x for x in options_5 if x <= A and A - x <= 15]:
                            for C in [x for x in options_5 if x <= B and B - x <= 15]:
                                for D in [x for x in options_5 if x <= C and C - x <= 15]:
                                    for E in [x for x in options_5 if x <= D and D - x <= 15]:
                                        for BOS in [50, 55, 60, 65, 70, 75, 80, 85, 90, 95]:
                                            combos.append((A, B, C, D, E, BOS))

                    results = []
                    for A, B, C, D, E, BOS in combos:
                        levels = {'A': A/100.0, 'B': B/100.0, 'C': C/100.0, 'D': D/100.0, 'E': E/100.0, 'BOŞ': BOS/100.0}
                        total_b = 0
                        for cls_k, val_v in levels.items():
                            if cls_k in grouped_stats:
                                z_new_val = NormalDist().inv_cdf(val_v)
                                total_b += grouped_stats[cls_k]['sum_min_z'] * z_new_val + grouped_stats[cls_k]['sum_fark']
                        results.append({'budget': total_b, 'A': A, 'B': B, 'C': C, 'D': D, 'E': E, 'BOS': BOS})

                    valid_res = [r for r in results if r['budget'] <= analiz_butce * 1.02]
                    valid_res.sort(key=lambda x: x['budget'])

                    if not valid_res:
                        valid_res = sorted(results, key=lambda x: x['budget'])[:5]

                    n_res = len(valid_res)
                    p_indices = [
                        ("1️⃣ Tasarruf Odaklı (Düşük Bütçe)", 0),
                        ("2️⃣ Ekonomik Dengeli", int(n_res * 0.25)),
                        ("3️⃣ Standart Optimal Seviye", int(n_res * 0.50)),
                        ("4️⃣ Yüksek Hizmet Seviyesi", int(n_res * 0.75)),
                        ("5️⃣ Bütçe Sınırına En Yakın Limit", n_res - 1)
                    ]

                    st.markdown("#### 🎯 Analiz Edilen Örnek Pareto Hizmet Senaryoları")
                    s_cols = st.columns(5)
                    for idx_i, (s_title, s_idx) in enumerate(p_indices):
                        s_idx = min(s_idx, n_res - 1)
                        scen = valid_res[s_idx]
                        
                        with s_cols[idx_i]:
                            st.subheader(s_title)
                            st.metric("Tahsili Bütçe", f"{scen['budget']:,.0f} ₺".replace(",", "."))
                            st.markdown(f"""
                            - **A Grubu:** %{scen['A']}
                            - **B Grubu:** %{scen['B']}
                            - **C Grubu:** %{scen['C']}
                            - **D Grubu:** %{scen['D']}
                            - **E Grubu:** %{scen['E']}
                            - **Boş/Tanımsız:** %{scen['BOS']}
                            """)
                            
                            if st.button(f"✅ Bu Senaryoyu Seç", key=f"btn_scen_{idx_i}"):
                                st.session_state["level_A"] = scen['A']
                                st.session_state["level_B"] = scen['B']
                                st.session_state["level_C"] = scen['C']
                                st.session_state["level_D"] = scen['D']
                                st.session_state["level_E"] = scen['E']
                                st.session_state["level_BOS"] = scen['BOS']
                                st.session_state["hedef_butce"] = int(analiz_butce)
                                st.session_state["form_submitted"] = True
                                st.rerun()

            # --- 4. HEDEF BÜTÇE VE PARETO SEÇİM FORMU (OKEY / HESAPLA BUTONLU) ---
            options_5 = list(range(50, 100, 5))

            with st.form(key="butce_hesaplama_formu"):
                st.markdown("### 🎯 2. Hedef Bütçe & Emniyet Seviyesi Ayarları")
                
                col_b1, col_b2 = st.columns([1, 1])
                with col_b1:
                    hedef_butce_input = st.number_input(
                        "Hedeflemek İstediğiniz Toplam Bütçe Sınırı (TL):",
                        value=int(st.session_state["hedef_butce"]),
                        step=50000000,
                        format="%d",
                        help="Örneğin 6.200.000.000 TL (6.2 Milyar) yazabilirsiniz."
                    )
                
                st.markdown("**Defa (Pareto) Sınıflarına Göre Emniyet Seviyeleri Seçimi:**")
                
                p1, p2, p3, p4, p5, p6 = st.columns(6)
                with p1:
                    idx_A = options_5.index(st.session_state["level_A"]) if st.session_state["level_A"] in options_5 else 8
                    level_A = st.selectbox("A Grubu Emniyet %", options=options_5, index=idx_A)
                with p2:
                    idx_B = options_5.index(st.session_state["level_B"]) if st.session_state["level_B"] in options_5 else 7
                    level_B = st.selectbox("B Grubu Emniyet %", options=options_5, index=idx_B)
                with p3:
                    idx_C = options_5.index(st.session_state["level_C"]) if st.session_state["level_C"] in options_5 else 6
                    level_C = st.selectbox("C Grubu Emniyet %", options=options_5, index=idx_C)
                with p4:
                    idx_D = options_5.index(st.session_state["level_D"]) if st.session_state["level_D"] in options_5 else 5
                    level_D = st.selectbox("D Grubu Emniyet %", options=options_5, index=idx_D)
                with p5:
                    idx_E = options_5.index(st.session_state["level_E"]) if st.session_state["level_E"] in options_5 else 4
                    level_E = st.selectbox("E Grubu Emniyet %", options=options_5, index=idx_E)
                with p6:
                    idx_BOS = options_5.index(st.session_state["level_BOS"]) if st.session_state["level_BOS"] in options_5 else 0
                    level_BOS = st.selectbox("Boş/Tanımsız %", options=options_5, index=idx_BOS, help="Defa sınıfı boş olan ürünler için emniyet yüzdesi")
                
                st.markdown("<br>", unsafe_allow_html=True)
                form_submitted = st.form_submit_button("🚀 Değerleri Onayla ve Hesapla", use_container_width=True)

            # Form gönderildiğinde session state güncelle
            if form_submitted:
                st.session_state["level_A"] = level_A
                st.session_state["level_B"] = level_B
                st.session_state["level_C"] = level_C
                st.session_state["level_D"] = level_D
                st.session_state["level_E"] = level_E
                st.session_state["level_BOS"] = level_BOS
                st.session_state["hedef_butce"] = hedef_butce_input
                st.session_state["form_submitted"] = True

            # --- 5. HESAPLAMA VE SONUÇ EKRANI ---
            if st.session_state["form_submitted"]:
                level_map = {
                    'A': st.session_state["level_A"],
                    'B': st.session_state["level_B"],
                    'C': st.session_state["level_C"],
                    'D': st.session_state["level_D"],
                    'E': st.session_state["level_E"],
                    'BOŞ': st.session_state["level_BOS"]
                }
                
                df['Yeni_Emniyet_Seviyesi'] = df['Temiz_Pareto'].map(level_map).fillna(st.session_state["level_BOS"])
                
                # NORMSINV Formülü ile Hesaplama
                old_levels = df['Emniyet Seviyesi'].clip(50, 99) / 100.0
                new_levels = df['Yeni_Emniyet_Seviyesi'] / 100.0
                
                z_old = old_levels.apply(lambda p: NormalDist().inv_cdf(p))
                z_old = np.where(z_old == 0, 1e-5, z_old)
                z_new = new_levels.apply(lambda p: NormalDist().inv_cdf(p))
                
                df['Hedef_Min_Fark'] = df['Hedef TL'] - df['Min TL']
                df['Yeni Min TL'] = df['Min TL'] * (z_new / z_old)
                df['Yeni Hedef TL'] = df['Yeni Min TL'] + df['Hedef_Min_Fark']
                df['Yeni Optimum TL'] = (df['Yeni Min TL'] + df['Yeni Hedef TL']) / 2.0
                df['Yeni Fazla TL'] = np.maximum(0, df['Stok TL'] - df['Yeni Hedef TL'])
                
                yeni_toplam_hedef = df['Yeni Hedef TL'].sum()
                yeni_toplam_fazla = df['Yeni Fazla TL'].sum()
                
                # --- 6. SONUÇ PANELLERİ ---
                st.markdown("---")
                st.markdown("### 📈 3. Yeni Belirlenen Hedef Seviyeleri & Bütçe Analizi")
                
                aktuel_butce = st.session_state["hedef_butce"]
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Yeni Hesaplanan Hedef TL", f"{yeni_toplam_hedef:,.0f} ₺".replace(",", "."))
                k2.metric("Girdiğiniz Bütçe Sınırı", f"{aktuel_butce:,.0f} ₺".replace(",", "."))
                
                fark = aktuel_butce - yeni_toplam_hedef
                if fark >= 0:
                    k3.metric("Kalan Bütçe Payı", f"{fark:,.0f} ₺".replace(",", "."))
                    st.success(f"🟢 **BÜTÇE UYGUN:** Belirlediğiniz emniyet seviyeleri ile oluşan Yeni Hedef TL ({yeni_toplam_hedef:,.0f} ₺), {aktuel_butce:,.0f} ₺ bütçe sınırınızın altında kalmaktadır.")
                else:
                    k3.metric("Bütçe Aşım Miktarı", f"{abs(fark):,.0f} ₺".replace(",", "."))
                    st.error(f"🔴 **BÜTÇE AŞILDI:** Oluşan Yeni Hedef TL ({yeni_toplam_hedef:,.0f} ₺), girdiğiniz {aktuel_butce:,.0f} ₺ sınırını {abs(fark):,.0f} ₺ aşıyor. Lütfen emniyet seviyelerini düşürün.")
                    
                k4.metric("Yeni Oluşan Fazla (Atıl) Stok TL", f"{yeni_toplam_fazla:,.0f} ₺".replace(",", "."))
                
                # --- 7. TABLO GÖSTERİMİ ---
                st.markdown("### 📋 4. Ürün Bazlı Detay Tablosu")
                
                gosterim_sutunlari = [
                    'Ürün Kodu', 'Ürün', 'Firma', pareto_col_name, 'Emniyet Seviyesi', 
                    'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL',
                    'Yeni_Emniyet_Seviyesi', 'Yeni Min TL', 'Yeni Optimum TL', 'Yeni Hedef TL', 'Yeni Fazla TL'
                ]
                gosterim_sutunlari = [c for c in gosterim_sutunlari if c in df.columns or c == pareto_col_name]
                
                st.dataframe(df[gosterim_sutunlari], use_container_width=True)
                
                # --- 8. EXCEL İNDİRME ---
                st.markdown("### 📥 5. Sonuçları Excel (.xlsx) Formatında İndirin")
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df[gosterim_sutunlari].to_excel(writer, index=False, sheet_name='Emniyet_Optimizasyonu')
                    
                st.download_button(
                    label="🟢 Sonuçları Excel (.xlsx) Olarak İndir",
                    data=excel_buffer.getvalue(),
                    file_name=f"Emniyet_Seviyeleri_Optimizasyon_{aktuel_butce}_TL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Dosya işlenirken bir hata oluştu: {e}")
