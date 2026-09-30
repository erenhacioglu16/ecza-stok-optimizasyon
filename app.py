import streamlit as st
import pandas as pd
import numpy as np
from statistics import NormalDist
import io

# --- SAYFA YAPILANDIRMASI VE TÜRKÇE FORMAT FONKSİYONU ---
st.set_page_config(page_title="Stok & Bütçe Optimizasyon Paneli", layout="wide", initial_sidebar_state="collapsed")

# KPI'lar için düz (küsüratsız) format
def format_tr(val):
    if pd.isna(val) or val == 0:
        return "0 ₺"
    s = f"{val:,.0f}"
    return s.replace(",", ".") + " ₺"

# Tablo içi için küsüratlı, noktalı-virgüllü format (Örn: 1.234.567,89)
def table_format_tr(val):
    if pd.isna(val): return ""
    try:
        s = f"{val:,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return val

# %99 Dahil Edilmiş Genişletilmiş Seçenekler
options_ui = [50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 99]

def safe_index(val, default=90):
    return options_ui.index(val) if val in options_ui else options_ui.index(default)

# Session State Varsayılan Değerlerini İlklendirme
default_states = {
    "level_A": 90, "level_B": 85, "level_C": 80, "level_D": 75, "level_E": 70, "level_BOS": 99,
    "hedef_butce": 6200000000
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
            
            # Sütun temizliği
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(subset=['Ürün Kodu']).copy()
            
            numeric_cols = ['Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Pareto Sütun tespiti ve Temizliği
            pareto_col_name = 'Defa Pareto' if 'Defa Pareto' in df.columns else ('Defa ABC' if 'Defa ABC' in df.columns else None)
            if pareto_col_name:
                pareto_series = df[pareto_col_name].fillna('BOŞ').astype(str).str.strip().str.upper()
                pareto_series = pareto_series.replace(['NAN', 'NONE', ''], 'BOŞ')
            else:
                pareto_series = pd.Series(['BOŞ'] * len(df), index=df.index)
            df['Temiz_Pareto'] = pareto_series

            # Ortak Ön Hesaplamalar
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

            st.markdown("---")

            # --- 3. AKILLI BÜTÇE & PARETO ÖNERİ ASİSTANI ---
            with st.expander("💡 🤖 Bütçeye Göre Akıllı Pareto Öneri Asistanı (Kademeli & Bütçe Aşım Analizi)", expanded=False):
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
                    options_no_50 = [55, 60, 65, 70, 75, 80, 85, 90, 95, 99]
                    
                    combos = [{'A': 50, 'B': 50, 'C': 50, 'D': 50, 'E': 50, 'BOS': 50, 'is_min': True}]
                    
                    for A in options_no_50:
                        for B in [x for x in options_no_50 if x <= A and A - x <= 20]:
                            for C in [x for x in options_no_50 if x <= B and B - x <= 20]:
                                for D in [x for x in options_no_50 if x <= C and C - x <= 20]:
                                    for E in [x for x in options_no_50 if x <= D and D - x <= 20]:
                                        combos.append({'A': A, 'B': B, 'C': C, 'D': D, 'E': E, 'BOS': 99, 'is_min': False})

                    results = []
                    for c_dict in combos:
                        z_vals = [z_dict[c_dict['A']], z_dict[c_dict['B']], z_dict[c_dict['C']], z_dict[c_dict['D']], z_dict[c_dict['E']], z_dict[c_dict['BOS']]]
                        total_opt = sum([np.sum(min_div_z_m[i] * z_vals[i] + fark_m[i] / 2.0) for i in range(6)])
                        total_fazla = sum([np.sum(np.maximum(0, stok_m[i] - (min_div_z_m[i] * z_vals[i] + fark_m[i]))) for i in range(6)])
                        
                        results.append({
                            'budget': total_opt + total_fazla, 'A': c_dict['A'], 'B': c_dict['B'], 'C': c_dict['C'], 
                            'D': c_dict['D'], 'E': c_dict['E'], 'BOS': c_dict['BOS'], 'is_min': c_dict['is_min']
                        })

                    min_b = min([r['budget'] for r in results])
                    max_possible = max([r['budget'] for r in results])
                    actual_target = min(analiz_butce, max_possible)
                    
                    target_points = [
                        min_b, min_b + (actual_target - min_b) * 0.40, min_b + (actual_target - min_b) * 0.65, 
                        min_b + (actual_target - min_b) * 0.85, actual_target, actual_target * 1.02, 
                        actual_target * 1.05, actual_target * 1.10
                    ]
                    
                    selected_scenarios = [[r for r in results if r['is_min']][0]]
                    seen_keys = {(50,50,50,50,50,50)}
                    
                    def smooth_score(cand):
                        diffs = [cand['A']-cand['B'], cand['B']-cand['C'], cand['C']-cand['D'], cand['D']-cand['E']]
                        return np.var(diffs)

                    for tp_idx, tp in enumerate(target_points[1:], start=1):
                        if tp > max_possible: tp = max_possible
                        sorted_by_dist = sorted([r for r in results if not r['is_min']], key=lambda x: abs(x['budget'] - tp))
                        budget_constraint = (actual_target * 1.01) if tp_idx < 5 else (actual_target * 1.15)
                        
                        valid_candidates = [c for c in sorted_by_dist if c['budget'] <= budget_constraint and (c['A'], c['B'], c['C'], c['D'], c['E'], c['BOS']) not in seen_keys][:30]
                        if valid_candidates:
                            best_smooth = sorted(valid_candidates, key=smooth_score)[0]
                            selected_scenarios.append(best_smooth)
                            seen_keys.add((best_smooth['A'], best_smooth['B'], best_smooth['C'], best_smooth['D'], best_smooth['E'], best_smooth['BOS']))

                    p_indices = ["1️⃣ Tasarruf (Minimum)", "2️⃣ Ekonomik Alt", "3️⃣ Ekonomik Üst", "4️⃣ Yüksek Hizmet", "5️⃣ Bütçe Sınırı (Hedef)", "6️⃣ Bütçe Aşımı (%2)", "7️⃣ Bütçe Aşımı (%5)", "8️⃣ Max Fırsat (%10)"]

                    st.markdown("#### 🎯 Analiz Edilen 8 Farklı Kademeli Pareto Senaryosu")
                    
                    st.markdown("##### 🟢 Hedef Bütçe İçi Senaryolar (Boş Gruplar %99 Koruma)")
                    row1 = st.columns(4)
                    for idx_i in range(4):
                        if idx_i >= len(selected_scenarios): continue
                        scen = selected_scenarios[idx_i]
                        with row1[idx_i]:
                            st.subheader(p_indices[idx_i])
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            st.metric("🎯 Yeni Hedef Değer", format_tr(scen['budget']))
                            if st.button(f"✅ Bu Senaryoyu Seç", key=f"btn_scen_{idx_i}"):
                                st.session_state.update({"level_A": scen['A'], "level_B": scen['B'], "level_C": scen['C'], "level_D": scen['D'], "level_E": scen['E'], "level_BOS": scen['BOS'], "hedef_butce": int(analiz_butce)})
                                st.rerun()

                    st.markdown("##### 🔴 Bütçe Limitine Yakın ve Limit Aşımı (Fırsat) Senaryoları")
                    row2 = st.columns(4)
                    for idx_i in range(4, 8):
                        if idx_i >= len(selected_scenarios): continue
                        scen = selected_scenarios[idx_i]
                        with row2[idx_i - 4]:
                            st.subheader(p_indices[idx_i])
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            delta_val = float(scen['budget']) - float(analiz_butce)
                            st.metric("🎯 Yeni Hedef Değer", format_tr(scen['budget']), delta=f"{format_tr(abs(delta_val))} Fark", delta_color="inverse" if delta_val > 0 else "normal")
                            if st.button(f"✅ Bu Senaryoyu Seç", key=f"btn_scen_{idx_i}"):
                                st.session_state.update({"level_A": scen['A'], "level_B": scen['B'], "level_C": scen['C'], "level_D": scen['D'], "level_E": scen['E'], "level_BOS": scen['BOS'], "hedef_butce": int(analiz_butce)})
                                st.rerun()

            # --- 4. HEDEF BÜTÇE VE PARETO SEÇİM FORMU (ANINDA TETİKLENİR) ---
            st.markdown("### 🛠️ 2. Stratejik Planlama: Hizmet Seviyelerini Belirle")
            
            col_b1, col_b2 = st.columns([1, 1])
            with col_b1:
                ui_butce = st.number_input(
                    "Yönetim Bütçe Sınırı / Hedef (TL):",
                    value=int(st.session_state["hedef_butce"]), step=50000000, format="%d"
                )
            
            st.markdown("**ABC/Pareto Sınıflarına Göre Hizmet (Emniyet) Seviyeleri (%):**")
            p1, p2, p3, p4, p5, p6 = st.columns(6)
            with p1: ui_A = st.selectbox("A Sınıfı", options_ui, index=safe_index(st.session_state["level_A"]))
            with p2: ui_B = st.selectbox("B Sınıfı", options_ui, index=safe_index(st.session_state["level_B"]))
            with p3: ui_C = st.selectbox("C Sınıfı", options_ui, index=safe_index(st.session_state["level_C"]))
            with p4: ui_D = st.selectbox("D Sınıfı", options_ui, index=safe_index(st.session_state["level_D"]))
            with p5: ui_E = st.selectbox("E Sınıfı", options_ui, index=safe_index(st.session_state["level_E"]))
            with p6: ui_BOS = st.selectbox("Tanımsız Sınıf", options_ui, index=safe_index(st.session_state["level_BOS"]))
            
            if st.button("🚀 Senaryoyu Hesapla ve Uygula", use_container_width=True):
                st.session_state.update({"level_A": ui_A, "level_B": ui_B, "level_C": ui_C, "level_D": ui_D, "level_E": ui_E, "level_BOS": ui_BOS, "hedef_butce": ui_butce})

            # --- 5. HESAPLAMA VE SONUÇ EKRANI (HER ZAMAN GÜNCEL) ---
            # Session state'deki güncel değerleri kullanarak hesapla
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
            
            # --- 7. TABLO GÖSTERİMİ VE İNDİRME ---
            st.markdown("### 📋 4. Ürün Bazlı Operasyonel Detay Tablosu")
            
            gosterim_sutunlari = [
                'Ürün Kodu', 'Ürün', pareto_col_name, 'Emniyet Seviyesi', 
                'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL',
                'Yeni_Emniyet_Seviyesi', 'Yeni Min TL', 'Yeni Optimum TL', 'Yeni Hedef TL', 'Yeni Fazla TL'
            ]
            gosterim_sutunlari = [c for c in gosterim_sutunlari if c in df.columns or c == pareto_col_name]
            
            # Tablo gösterimi öncesi sütunları noktalı-virgüllü formata çevirme
            df_gosterim = df[gosterim_sutunlari].copy()
            for col in ['Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL', 'Yeni Min TL', 'Yeni Optimum TL', 'Yeni Hedef TL', 'Yeni Fazla TL']:
                if col in df_gosterim.columns:
                    df_gosterim[col] = df_gosterim[col].apply(table_format_tr)
            
            st.dataframe(df_gosterim, use_container_width=True)
            
            st.markdown("### 📥 5. Veri Dışa Aktarımı")
            st.info("Aşağıdaki yeşil butona tıklayarak, içerisinde Pareto sınıflarının ve yeni/eski emniyet seviyelerinin olduğu gerçek bir Excel (**.xlsx**) dosyası indirebilirsiniz.")
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # İndirilecek excel ham sayı formatında olmalı ki sonrasında işlem yapılabilsin
                df[gosterim_sutunlari].to_excel(writer, index=False, sheet_name='Stok_Optimizasyon_Sonucu')
                
            st.download_button(
                label="🟢 Yeni Stok Senaryosunu Excel (.xlsx) Olarak İndir", data=excel_buffer.getvalue(),
                file_name=f"Stok_Optimizasyon_{aktuel_butce}_TL.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Dosya işlenirken bir hata oluştu: {e}")
