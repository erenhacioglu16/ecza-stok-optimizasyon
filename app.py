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
    s = f"{val:,.0f}"
    return s.replace(",", ".") + " ₺"

# %99 Dahil Edilmiş Genişletilmiş Seçenekler
options_ui = [50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 99]

def safe_index(val, default=90):
    return options_ui.index(val) if val in options_ui else options_ui.index(default)

# Session State Varsayılan Değerlerini İlklendirme
default_states = {
    "level_A": 90, "level_B": 85, "level_C": 80, "level_D": 75, "level_E": 70, "level_BOS": 99,
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
    st.markdown("Stok parametrelerinizi analiz edin, kademeli hizmet seviyesi senaryoları kurgulayın ve bütçenizi kontrol altına alın.")
    
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

            # Ortak Z-Score ve Farkların Ön Hesaplaması
            old_levels_all = df['Emniyet Seviyesi'].clip(50, 99) / 100.0
            z_old_all = old_levels_all.apply(lambda p: NormalDist().inv_cdf(p))
            z_old_all = np.where(z_old_all == 0, 1e-5, z_old_all)
            df['Hedef_Min_Fark'] = df['Hedef TL'] - df['Min TL']
            
            st.success(f"✅ **{uploaded_file.name}** veritabanına başarıyla aktarıldı.")
            
            # ---------------------------------------------------------
            # YÖNETİCİ ÖZETİ - GÜNCEL DURUM
            # ---------------------------------------------------------
            st.markdown("---")
            st.markdown("### 📌 Yönetici Özeti: Mevcut Envanter Durumu")
            
            mevcut_optimum_toplam = df['Optimum TL'].sum()
            mevcut_fazla_toplam = df['Fazla TL'].sum()
            mevcut_stok_toplam = df['Stok TL'].sum()
            mevcut_hedef_toplam = df['Hedef TL'].sum()
            mevcut_min_toplam = df['Min TL'].sum()
            
            guncel_hedef_deger = mevcut_optimum_toplam + mevcut_fazla_toplam
            stok_hedef_farki = mevcut_stok_toplam - guncel_hedef_deger
            
            kpi1, kpi2, kpi3 = st.columns(3)
            with kpi1:
                st.info(f"**Güncel Hedef Değer**\n### {format_tr(guncel_hedef_deger)}\n*(Optimum TL + Fazla TL)*")
            with kpi2:
                st.info(f"**Stok İle Hedef Değer Farkı**\n### {format_tr(stok_hedef_farki)}\n*(Stok TL - Güncel Hedef Değer)*")
            with kpi3:
                st.info(f"**Mevcut Stok Değeri (Fiili)**\n### {format_tr(mevcut_stok_toplam)}\n*(Depodaki Toplam Stok)*")

            st.markdown("<br>", unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Mevcut Min TL", format_tr(mevcut_min_toplam))
            m2.metric("Mevcut Optimum TL", format_tr(mevcut_optimum_toplam))
            m3.metric("Mevcut Hedef TL (Sistem)", format_tr(mevcut_hedef_toplam))
            m4.metric("Mevcut Fazla/Atıl TL", format_tr(mevcut_fazla_toplam))
            st.markdown("---")

            # --- 3. AKILLI BÜTÇE & PARETO ÖNERİ ASİSTANI ---
            with st.expander("💡 🤖 Bütçeye Göre Akıllı Pareto Öneri Asistanı (Kademeli & Bütçe Aşım Analizi)", expanded=False):
                st.markdown("Sistem bütçenizi tarayarak A'dan E'ye **düzenli kademeler halinde düşen (pürüzsüz)** senaryoları seçer. Tanımsız sınıflar %99 olarak korumaya alınır.")
                
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
                    run_analysis = st.button("🔍 Kademeli 8 Senaryoyu Analiz Et ve Getir", use_container_width=True)

                if run_analysis:
                    min_div_z_arr = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] / z_old_all)
                    fark_arr = df['Hedef_Min_Fark'].values
                    stok_arr = df['Stok TL'].values
                    
                    cls_map = {'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'BOŞ':5}
                    cls_idx = df['Temiz_Pareto'].map(cls_map).fillna(5).astype(int).values
                    
                    masks = [cls_idx == i for i in range(6)]
                    min_div_z_m = [min_div_z_arr[m] for m in masks]
                    fark_m = [fark_arr[m] for m in masks]
                    stok_m = [stok_arr[m] for m in masks]
                    
                    z_dict = {val: NormalDist().inv_cdf(val/100.0) for val in options_ui}
                    
                    # KURALLAR:
                    # - 50 sadece 1. Senaryoda var, diğerlerinde yok (55'ten başlıyor).
                    # - BOS her zaman 99 (1. Senaryo hariç).
                    options_no_50 = [55, 60, 65, 70, 75, 80, 85, 90, 95, 99]
                    
                    combos = []
                    # Sadece 1 adet "Tasarruf Odaklı (Min)" senaryosu için en dip seviyeyi özel olarak ekliyoruz:
                    combos.append({'A': 50, 'B': 50, 'C': 50, 'D': 50, 'E': 50, 'BOS': 50, 'is_min': True})
                    
                    for A in options_no_50:
                        for B in [x for x in options_no_50 if x <= A and A - x <= 20]:
                            for C in [x for x in options_no_50 if x <= B and B - x <= 20]:
                                for D in [x for x in options_no_50 if x <= C and C - x <= 20]:
                                    for E in [x for x in options_no_50 if x <= D and D - x <= 20]:
                                        combos.append({'A': A, 'B': B, 'C': C, 'D': D, 'E': E, 'BOS': 99, 'is_min': False})

                    results = []
                    for c_dict in combos:
                        A, B, C, D, E, BOS = c_dict['A'], c_dict['B'], c_dict['C'], c_dict['D'], c_dict['E'], c_dict['BOS']
                        z_vals = [z_dict[A], z_dict[B], z_dict[C], z_dict[D], z_dict[E], z_dict[BOS]]
                        total_opt = 0
                        total_fazla = 0
                        
                        for i in range(6):
                            y_min = min_div_z_m[i] * z_vals[i]
                            y_hedef = y_min + fark_m[i]
                            y_opt = y_min + fark_m[i] / 2.0
                            total_opt += np.sum(y_opt)
                            total_fazla += np.sum(np.maximum(0, stok_m[i] - y_hedef))
                        
                        target_budget = total_opt + total_fazla
                        results.append({
                            'budget': target_budget, 'A': A, 'B': B, 'C': C, 'D': D, 'E': E, 'BOS': BOS, 
                            'is_min': c_dict['is_min']
                        })

                    # Hedef Noktaları
                    min_b = min([r['budget'] for r in results])
                    max_possible = max([r['budget'] for r in results])
                    actual_target = min(analiz_butce, max_possible)
                    
                    target_points = [
                        min_b,                                         # 1. En düşük (Tasarruf)
                        min_b + (actual_target - min_b) * 0.40,        # 2. Ekonomik Alt
                        min_b + (actual_target - min_b) * 0.65,        # 3. Ekonomik Üst
                        min_b + (actual_target - min_b) * 0.85,        # 4. Yüksek Hizmet
                        actual_target,                                 # 5. Tam Sınır (İstenen Bütçe)
                        actual_target * 1.02,                          # 6. %2 Bütçe Aşımı
                        actual_target * 1.05,                          # 7. %5 Bütçe Aşımı
                        actual_target * 1.10                           # 8. %10 Bütçe Aşımı
                    ]
                    
                    selected_scenarios = []
                    seen_keys = set()
                    
                    # 1. Tasarruf (Zorunlu En Düşük)
                    min_scen = [r for r in results if r['is_min']][0]
                    selected_scenarios.append(min_scen)
                    seen_keys.add((50,50,50,50,50,50))
                    
                    # Kademeli Şelale (Smoothness) Puanlayıcı Fonksiyon
                    # A-B, B-C, C-D, D-E farklarının varyansı (değişkenliği) ne kadar azsa şelale o kadar pürüzsüzdür.
                    def smooth_score(cand):
                        diffs = [cand['A']-cand['B'], cand['B']-cand['C'], cand['C']-cand['D'], cand['D']-cand['E']]
                        return np.var(diffs)

                    for tp_idx, tp in enumerate(target_points):
                        if tp_idx == 0: continue # 1.yi zaten aldık
                        if tp > max_possible: tp = max_possible
                            
                        # Bu hedefe bütçe olarak yakın olan en iyi 30 senaryoyu çek
                        sorted_by_dist = sorted([r for r in results if not r['is_min']], key=lambda x: abs(x['budget'] - tp))
                        budget_constraint = (actual_target * 1.01) if tp_idx < 5 else (actual_target * 1.15)
                        
                        valid_candidates = [c for c in sorted_by_dist if c['budget'] <= budget_constraint and (c['A'], c['B'], c['C'], c['D'], c['E'], c['BOS']) not in seen_keys][:30]
                        
                        if valid_candidates:
                            # Yakın olanlar arasından 'en düzenli / kademeli' inen senaryoyu seç
                            best_smooth = sorted(valid_candidates, key=smooth_score)[0]
                            selected_scenarios.append(best_smooth)
                            seen_keys.add((best_smooth['A'], best_smooth['B'], best_smooth['C'], best_smooth['D'], best_smooth['E'], best_smooth['BOS']))

                    p_indices = [
                        "1️⃣ Tasarruf (Minimum)", "2️⃣ Ekonomik Alt", "3️⃣ Ekonomik Üst", "4️⃣ Yüksek Hizmet",
                        "5️⃣ Bütçe Sınırı (Hedef)", "6️⃣ Bütçe Aşımı (%2)", "7️⃣ Bütçe Aşımı (%5)", "8️⃣ Max Fırsat (%10)"
                    ]

                    st.markdown("#### 🎯 Analiz Edilen 8 Farklı Kademeli Pareto Senaryosu")
                    
                    st.markdown("##### 🟢 Hedef Bütçe İçi Senaryolar (Boş Gruplar %99 Koruma)")
                    row1 = st.columns(4)
                    for idx_i in range(4):
                        if idx_i >= len(selected_scenarios): continue
                        scen = selected_scenarios[idx_i]
                        
                        level_map_scen = {'A': scen['A'], 'B': scen['B'], 'C': scen['C'], 'D': scen['D'], 'E': scen['E'], 'BOŞ': scen['BOS']}
                        scen_emniyet = df['Temiz_Pareto'].map(level_map_scen).fillna(scen['BOS']) / 100.0
                        z_new_scen = scen_emniyet.apply(lambda p: NormalDist().inv_cdf(p))
                        
                        scen_yeni_min = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] * (z_new_scen / z_old_all))
                        scen_yeni_hedef = scen_yeni_min + df['Hedef_Min_Fark']
                        scen_yeni_optimum = (scen_yeni_min + scen_yeni_hedef) / 2.0
                        scen_yeni_fazla = np.maximum(0, df['Stok TL'] - scen_yeni_hedef)
                        scen_guncel_hedef = scen_yeni_optimum.sum() + scen_yeni_fazla.sum()
                        
                        with row1[idx_i]:
                            st.subheader(p_indices[idx_i])
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            st.metric("🎯 Yeni Hedef Değer", format_tr(scen_guncel_hedef))
                            st.metric("⚖️ Yeni Optimum TL", format_tr(scen_yeni_optimum.sum()))
                            st.metric("⚠️ Yeni Fazla TL", format_tr(scen_yeni_fazla.sum()))
                            
                            if st.button(f"✅ Bu Senaryoyu Seç", key=f"btn_scen_{idx_i}"):
                                st.session_state["level_A"], st.session_state["level_B"] = scen['A'], scen['B']
                                st.session_state["level_C"], st.session_state["level_D"] = scen['C'], scen['D']
                                st.session_state["level_E"], st.session_state["level_BOS"] = scen['E'], scen['BOS']
                                st.session_state["hedef_butce"] = int(analiz_butce)
                                st.session_state["form_submitted"] = True
                                st.rerun()

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown("##### 🔴 Bütçe Limitine Yakın ve Limit Aşımı (Fırsat) Senaryoları")
                    row2 = st.columns(4)
                    for idx_i in range(4, 8):
                        if idx_i >= len(selected_scenarios): continue
                        scen = selected_scenarios[idx_i]
                        
                        level_map_scen = {'A': scen['A'], 'B': scen['B'], 'C': scen['C'], 'D': scen['D'], 'E': scen['E'], 'BOŞ': scen['BOS']}
                        scen_emniyet = df['Temiz_Pareto'].map(level_map_scen).fillna(scen['BOS']) / 100.0
                        z_new_scen = scen_emniyet.apply(lambda p: NormalDist().inv_cdf(p))
                        
                        scen_yeni_min = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] * (z_new_scen / z_old_all))
                        scen_yeni_hedef = scen_yeni_min + df['Hedef_Min_Fark']
                        scen_yeni_optimum = (scen_yeni_min + scen_yeni_hedef) / 2.0
                        scen_yeni_fazla = np.maximum(0, df['Stok TL'] - scen_yeni_hedef)
                        scen_guncel_hedef = scen_yeni_optimum.sum() + scen_yeni_fazla.sum()
                        
                        with row2[idx_i - 4]:
                            st.subheader(p_indices[idx_i])
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            
                            delta_val = float(scen_guncel_hedef) - float(analiz_butce)
                            delta_color = "inverse" if delta_val > 0 else "normal"
                            
                            st.metric("🎯 Yeni Hedef Değer", format_tr(scen_guncel_hedef), delta=f"{format_tr(abs(delta_val))} Fark", delta_color=delta_color)
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
            with st.form(key="butce_hesaplama_formu"):
                st.markdown("### 🛠️ 2. Stratejik Planlama: Hizmet Seviyelerini Belirle")
                
                col_b1, col_b2 = st.columns([1, 1])
                with col_b1:
                    hedef_butce_input = st.number_input(
                        "Yönetim Bütçe Sınırı / Hedef (TL):",
                        value=int(st.session_state["hedef_butce"]), step=50000000, format="%d",
                        help="Analiz ve kurgu için Yeni Hedef Değer'i (Optimum + Fazla) baz alır."
                    )
                
                st.markdown("**ABC/Pareto Sınıflarına Göre Hizmet (Emniyet) Seviyeleri (%):**")
                p1, p2, p3, p4, p5, p6 = st.columns(6)
                with p1: level_A = st.selectbox("A Sınıfı", options_ui, index=safe_index(st.session_state.get("level_A", 90)))
                with p2: level_B = st.selectbox("B Sınıfı", options_ui, index=safe_index(st.session_state.get("level_B", 85)))
                with p3: level_C = st.selectbox("C Sınıfı", options_ui, index=safe_index(st.session_state.get("level_C", 80)))
                with p4: level_D = st.selectbox("D Sınıfı", options_ui, index=safe_index(st.session_state.get("level_D", 75)))
                with p5: level_E = st.selectbox("E Sınıfı", options_ui, index=safe_index(st.session_state.get("level_E", 70)))
                with p6: level_BOS = st.selectbox("Tanımsız Sınıf", options_ui, index=safe_index(st.session_state.get("level_BOS", 99)))
                
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
                
                yeni_toplam_optimum = df['Yeni Optimum TL'].sum()
                yeni_toplam_fazla = df['Yeni Fazla TL'].sum()
                
                yeni_hedef_deger = yeni_toplam_optimum + yeni_toplam_fazla
                yeni_stok_hedef_farki = mevcut_stok_toplam - yeni_hedef_deger
                yeni_toplam_hedef = df['Yeni Hedef TL'].sum()
                
                st.markdown("---")
                st.markdown("### 📈 3. Tedarik Zinciri & Senaryo Sonuçları")
                
                res1, res2, res3 = st.columns(3)
                with res1:
                    st.success(f"**YENİ Hedef Değer (Kurgulanan)**\n### {format_tr(yeni_hedef_deger)}\n*(Yeni Optimum TL + Yeni Fazla TL)*")
                with res2:
                    st.success(f"**YENİ Stok İle Hedef Değer Farkı**\n### {format_tr(yeni_stok_hedef_farki)}\n*(Stok TL - Yeni Hedef Değer)*")
                with res3:
                    st.success(f"**YENİ Sistem Hedef TL (Taban Limit)**\n### {format_tr(yeni_toplam_hedef)}\n*(Senaryoya ait sistem hedef tabanı)*")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                aktuel_butce = st.session_state["hedef_butce"]
                fark_butce = aktuel_butce - yeni_hedef_deger
                if fark_butce >= 0:
                    st.success(f"🟢 **BÜTÇE UYGUN:** Kurguladığınız **Yeni Hedef Değer ({format_tr(yeni_hedef_deger)})**, belirlediğiniz **{format_tr(aktuel_butce)}** sınırının altındadır. *(Avantaj: {format_tr(fark_butce)})*")
                else:
                    st.error(f"🔴 **BÜTÇE AŞILDI (OVER-BUDGET):** Kurguladığınız **Yeni Hedef Değer ({format_tr(yeni_hedef_deger)})**, belirlediğiniz **{format_tr(aktuel_butce)}** sınırını **{format_tr(abs(fark_butce))}** tutarında aşmaktadır!")
                
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Yeni Hedef TL (Sistem)", format_tr(yeni_toplam_hedef))
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
