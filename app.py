import streamlit as st
import pandas as pd
import numpy as np
from statistics import NormalDist
import io

# --- SAYFA YAPILANDIRMASI VE TÜRKÇE FORMAT FONKSİYONU ---
st.set_page_config(page_title="Stok & Bütçe Optimizasyon Paneli", layout="wide", initial_sidebar_state="collapsed")

def format_tr(val):
    if pd.isna(val) or val == 0:
        return "0 ₺"
    # Binlik ayraçları nokta yapmak için formatlama
    s = f"{val:,.0f}"
    return s.replace(",", ".") + " ₺"

# Session State Varsayılan Değerlerini İlklendirme
default_states = {
    "level_A": 90, "level_B": 85, "level_C": 80, "level_D": 75, "level_E": 70, "level_BOS": 50,
    "hedef_butce": 6200000000, "form_submitted": False
}
for key, val in default_states.items():
    if key not in st.session_state:
        st.session_state[key] = val

# --- 1. ŞİFRE EKRANI ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.title("🔐 Tedarik Zinciri Yönetimi - Güvenli Giriş")
        st.markdown("Devam etmek için lütfen yönetici şifrenizi giriniz.")
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
    st.title("📊 Envanter & Bütçe Optimizasyon Paneli")
    st.markdown("Stok parametrelerinizi analiz edin, yeni hizmet seviyesi senaryoları kurgulayın ve atıl bütçenizi kontrol altına alın.")
    
    # --- 2. DOSYA YÜKLEME ---
    st.markdown("### 📄 1. Veri Yükleme")
    uploaded_files = st.file_uploader(
        "'Emniyet Seviyesi' Excel dosyasını seçin", 
        type=["xlsx", "xls", "xlsb"], accept_multiple_files=True
    )
    
    if uploaded_files:
        uploaded_file = uploaded_files[-1]
        try:
            engine_choice = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else None
            df = pd.read_excel(uploaded_file, sheet_name='Stok seviyesi', skiprows=6, engine=engine_choice)
            
            # Sütun isimlerini ve sayısal verileri temizleme
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(subset=['Ürün Kodu']).copy()
            
            numeric_cols = ['Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL']
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

            # Ortak Z-Score (I Sütunu) ve Farkların Excel mantığıyla ön hesaplaması
            old_levels_all = df['Emniyet Seviyesi'].clip(50, 99) / 100.0
            z_old_all = old_levels_all.apply(lambda p: NormalDist().inv_cdf(p))
            z_old_all = np.where(z_old_all == 0, 1e-5, z_old_all)
            df['Hedef_Min_Fark'] = df['Hedef TL'] - df['Min TL']
            
            st.success(f"✅ **{uploaded_file.name}** veritabanına başarıyla aktarıldı.")
            
            # ---------------------------------------------------------
            # YÖNETİCİ ÖZETİ - GÜNCEL DURUM (KPI PANOLARI)
            # ---------------------------------------------------------
            st.markdown("---")
            st.markdown("### 📌 Yönetici Özeti: Mevcut Envanter Durumu")
            
            # Toplam hesaplamalar
            mevcut_optimum_toplam = df['Optimum TL'].sum()
            mevcut_fazla_toplam = df['Fazla TL'].sum()
            mevcut_stok_toplam = df['Stok TL'].sum()
            mevcut_hedef_toplam = df['Hedef TL'].sum()
            mevcut_min_toplam = df['Min TL'].sum()
            
            # Kritik KPI formülleri
            guncel_hedef_deger = mevcut_optimum_toplam + mevcut_fazla_toplam
            stok_hedef_farki = mevcut_stok_toplam - guncel_hedef_deger
            
            # 3'lü Kritik Gösterge
            kpi1, kpi2, kpi3 = st.columns(3)
            with kpi1:
                st.info(f"**Güncel Hedef Değer**\n### {format_tr(guncel_hedef_deger)}\n*(Optimum TL + Fazla TL)*")
            with kpi2:
                # Fark pozitifse stok şişkin, negatifse eksik
                renk = "error" if stok_hedef_farki > 0 else "success"
                st.info(f"**Stok İle Hedef Değer Farkı**\n### {format_tr(stok_hedef_farki)}\n*(Stok TL - Güncel Hedef Değer)*")
            with kpi3:
                st.info(f"**Mevcut Stok Değeri (Fiili)**\n### {format_tr(mevcut_stok_toplam)}\n*(Depodaki Toplam Stok)*")

            # Detay Metrikler
            st.markdown("<br>", unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Mevcut Min TL", format_tr(mevcut_min_toplam))
            m2.metric("Mevcut Optimum TL", format_tr(mevcut_optimum_toplam))
            m3.metric("Mevcut Hedef TL (Sistem)", format_tr(mevcut_hedef_toplam))
            m4.metric("Mevcut Fazla/Atıl TL", format_tr(mevcut_fazla_toplam))
            st.markdown("---")


            # --- 3. AKILLI BÜTÇE & PARETO ÖNERİ ASİSTANI ---
            with st.expander("💡 🤖 Bütçeye Göre Akıllı Pareto Öneri Asistanı (Otomatik Senaryo Analizi)", expanded=False):
                st.markdown("Sistem, tıpkı Excel'deki **R Sütunu** dinamikleriyle arka planda binlerce kombinasyon deneyerek size en uygun bütçe profillerini oluşturur.")
                
                col_rec1, col_rec2 = st.columns([2, 1])
                with col_rec1:
                    analiz_butce = st.number_input(
                        "Analiz Edilecek Hedef Bütçe Sınırı (TL):",
                        value=int(st.session_state["hedef_butce"]),
                        step=50000000, format="%d", key="analiz_butce_input"
                    )
                with col_rec2:
                    st.write("")
                    st.write("")
                    run_analysis = st.button("🔍 Senaryoları Analiz Et ve Getir", use_container_width=True)

                if run_analysis:
                    df_temp = pd.DataFrame({
                        'cls': df['Temiz_Pareto'],
                        'min_div_z': np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] / z_old_all),
                        'fark': df['Hedef_Min_Fark']
                    })
                    grouped_stats = df_temp.groupby('cls').agg(
                        sum_min_z=('min_div_z', 'sum'), sum_fark=('fark', 'sum')
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
                        ("1️⃣ Tasarruf Odaklı (Düşük)", 0),
                        ("2️⃣ Ekonomik Dengeli", int(n_res * 0.25)),
                        ("3️⃣ Standart Optimal", int(n_res * 0.50)),
                        ("4️⃣ Yüksek Hizmet", int(n_res * 0.75)),
                        ("5️⃣ Bütçeye En Yakın", n_res - 1)
                    ]

                    st.markdown("#### 🎯 Analiz Edilen Örnek Pareto Hizmet Senaryoları")
                    s_cols = st.columns(5)
                    
                    for idx_i, (s_title, s_idx) in enumerate(p_indices):
                        s_idx = min(s_idx, n_res - 1)
                        scen = valid_res[s_idx]
                        
                        level_map_scen = {'A': scen['A'], 'B': scen['B'], 'C': scen['C'], 'D': scen['D'], 'E': scen['E'], 'BOŞ': scen['BOS']}
                        scen_emniyet = df['Temiz_Pareto'].map(level_map_scen).fillna(scen['BOS']) / 100.0
                        z_new_scen = scen_emniyet.apply(lambda p: NormalDist().inv_cdf(p))
                        
                        scen_yeni_min = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] * (z_new_scen / z_old_all))
                        scen_yeni_hedef = scen_yeni_min + df['Hedef_Min_Fark']
                        scen_yeni_optimum = (scen_yeni_min + scen_yeni_hedef) / 2.0
                        scen_yeni_fazla = np.maximum(0, df['Stok TL'] - scen_yeni_hedef)
                        
                        # Senaryo İçi Tedarik Yöneticisi Özeti
                        scen_guncel_hedef = scen_yeni_optimum.sum() + scen_yeni_fazla.sum()
                        
                        with s_cols[idx_i]:
                            st.subheader(s_title)
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            st.metric("🎯 Yeni Hedef TL", format_tr(scen_yeni_hedef.sum()))
                            st.metric("📉 Yeni Min TL", format_tr(scen_yeni_min.sum()))
                            st.metric("⚖️ Yeni Optimum TL", format_tr(scen_yeni_optimum.sum()))
                            st.metric("⚠️ Yeni Fazla TL", format_tr(scen_yeni_fazla.sum()))
                            
                            if st.button(f"✅ Bu Senaryoyu Seç", key=f"btn_scen_{idx_i}"):
                                st.session_state["level_A"], st.session_state["level_B"] = scen['A'], scen['B']
                                st.session_state["level_C"], st.session_state["level_D"] = scen['C'], scen['D']
                                st.session_state["level_E"], st.session_state["level_BOS"] = scen['E'], scen['BOS']
                                st.session_state["hedef_butce"] = int(analiz_butce)
                                st.session_state["form_submitted"] = True
                                st.rerun()

            # --- 4. HEDEF BÜTÇE VE PARETO SEÇİM FORMU ---
            options_5 = list(range(50, 100, 5))
            with st.form(key="butce_hesaplama_formu"):
                st.markdown("### 🛠️ 2. Stratejik Planlama: Hizmet Seviyelerini Belirle")
                
                col_b1, col_b2 = st.columns([1, 1])
                with col_b1:
                    hedef_butce_input = st.number_input(
                        "Yönetim Bütçe Sınırı / Hedef (TL):",
                        value=int(st.session_state["hedef_butce"]), step=50000000, format="%d"
                    )
                
                st.markdown("**ABC/Pareto Sınıflarına Göre Hizmet (Emniyet) Seviyeleri (%):**")
                p1, p2, p3, p4, p5, p6 = st.columns(6)
                with p1: level_A = st.selectbox("A Sınıfı", options_5, index=options_5.index(st.session_state.get("level_A", 90)))
                with p2: level_B = st.selectbox("B Sınıfı", options_5, index=options_5.index(st.session_state.get("level_B", 85)))
                with p3: level_C = st.selectbox("C Sınıfı", options_5, index=options_5.index(st.session_state.get("level_C", 80)))
                with p4: level_D = st.selectbox("D Sınıfı", options_5, index=options_5.index(st.session_state.get("level_D", 75)))
                with p5: level_E = st.selectbox("E Sınıfı", options_5, index=options_5.index(st.session_state.get("level_E", 70)))
                with p6: level_BOS = st.selectbox("Tanımsız Sınıf", options_5, index=options_5.index(st.session_state.get("level_BOS", 50)))
                
                st.markdown("<br>", unsafe_allow_html=True)
                form_submitted = st.form_submit_button("🚀 Senaryoyu Hesapla ve Uygula", use_container_width=True)

            if form_submitted:
                st.session_state.update({
                    "level_A": level_A, "level_B": level_B, "level_C": level_C,
                    "level_D": level_D, "level_E": level_E, "level_BOS": level_BOS,
                    "hedef_butce": hedef_butce_input, "form_submitted": True
                })

            # --- 5. HESAPLAMA VE SONUÇ EKRANI (TEDARİK ZİNCİRİ ÖZETİ) ---
            if st.session_state["form_submitted"]:
                level_map = {
                    'A': st.session_state["level_A"], 'B': st.session_state["level_B"],
                    'C': st.session_state["level_C"], 'D': st.session_state["level_D"],
                    'E': st.session_state["level_E"], 'BOŞ': st.session_state["level_BOS"]
                }
                
                df['Yeni_Emniyet_Seviyesi'] = df['Temiz_Pareto'].map(level_map).fillna(st.session_state["level_BOS"])
                
                new_levels = df['Yeni_Emniyet_Seviyesi'] / 100.0
                z_new = new_levels.apply(lambda p: NormalDist().inv_cdf(p))
                
                df['Yeni Min TL'] = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] * (z_new / z_old_all))
                df['Yeni Hedef TL'] = df['Yeni Min TL'] + df['Hedef_Min_Fark']
                df['Yeni Optimum TL'] = (df['Yeni Min TL'] + df['Yeni Hedef TL']) / 2.0
                df['Yeni Fazla TL'] = np.maximum(0, df['Stok TL'] - df['Yeni Hedef TL'])
                
                # --- YENİ TEDARİK KPI HESAPLARI ---
                yeni_toplam_optimum = df['Yeni Optimum TL'].sum()
                yeni_toplam_fazla = df['Yeni Fazla TL'].sum()
                
                yeni_hedef_deger = yeni_toplam_optimum + yeni_toplam_fazla
                yeni_stok_hedef_farki = mevcut_stok_toplam - yeni_hedef_deger
                yeni_toplam_hedef = df['Yeni Hedef TL'].sum()
                
                st.markdown("---")
                st.markdown("### 📈 3. Tedarik Zinciri & Senaryo Sonuçları")
                
                # YENİ KPI KARTLARI
                res1, res2, res3 = st.columns(3)
                with res1:
                    st.success(f"**YENİ Hedef Değer (Kurgulanan)**\n### {format_tr(yeni_hedef_deger)}\n*(Yeni Optimum TL + Yeni Fazla TL)*")
                with res2:
                    st.success(f"**YENİ Stok İle Hedef Değer Farkı**\n### {format_tr(yeni_stok_hedef_farki)}\n*(Stok TL - Yeni Hedef Değer)*")
                with res3:
                    st.success(f"**YENİ Hedef TL Toplamı (Bütçe Limit Kontrolü)**\n### {format_tr(yeni_toplam_hedef)}\n*(Seçilen senaryoya ait taban hedef)*")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Detaylar ve Bütçe Uygunluğu
                aktuel_butce = st.session_state["hedef_butce"]
                fark_butce = aktuel_butce - yeni_toplam_hedef
                if fark_butce >= 0:
                    st.success(f"🟢 **BÜTÇE UYGUN:** Yeni oluşturulan **Hedef TL ({format_tr(yeni_toplam_hedef)})**, belirlediğiniz **{format_tr(aktuel_butce)}** bütçe sınırının altındadır. *(Avantaj: {format_tr(fark_butce)})*")
                else:
                    st.error(f"🔴 **BÜTÇE AŞILDI:** Yeni oluşturulan **Hedef TL ({format_tr(yeni_toplam_hedef)})**, belirlediğiniz **{format_tr(aktuel_butce)}** bütçe sınırını **{format_tr(abs(fark_butce))}** tutarında aşmaktadır! Hizmet seviyelerini düşürünüz.")
                
                # Sütun halinde alt metrikler
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Yeni Hedef TL (Detay)", format_tr(yeni_toplam_hedef))
                k2.metric("Yeni Min TL", format_tr(df['Yeni Min TL'].sum()))
                k3.metric("Yeni Optimum TL", format_tr(yeni_toplam_optimum))
                k4.metric("Yeni Fazla (Atıl) TL", format_tr(yeni_toplam_fazla))
                
                # --- 7. TABLO GÖSTERİMİ VE İNDİRME ---
                st.markdown("### 📋 4. Ürün Bazlı Operasyonel Detay Tablosu")
                
                gosterim_sutunlari = [
                    'Ürün Kodu', 'Ürün', pareto_col_name, 'Emniyet Seviyesi', 
                    'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL',
                    'Yeni_Emniyet_Seviyesi', 'Yeni Min TL', 'Yeni Optimum TL', 'Yeni Hedef TL', 'Yeni Fazla TL'
                ]
                gosterim_sutunlari = [c for c in gosterim_sutunlari if c in df.columns or c == pareto_col_name]
                
                st.dataframe(df[gosterim_sutunlari], use_container_width=True)
                
                st.markdown("### 📥 5. Veri Dışa Aktarımı")
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df[gosterim_sutunlari].to_excel(writer, index=False, sheet_name='Stok_Optimizasyon_Sonucu')
                    
                st.download_button(
                    label="🟢 Yeni Stok Senaryosunu Excel (.xlsx) Olarak İndir", data=excel_buffer.getvalue(),
                    file_name=f"Stok_Optimizasyon_{aktuel_butce}_TL.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Dosya işlenirken bir hata oluştu: {e}")
