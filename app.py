import streamlit as st
import pandas as pd
import numpy as np
from statistics import NormalDist
import io
import openpyxl

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(page_title="BEK Envanter & Bütçe Optimizasyon Paneli", layout="wide", initial_sidebar_state="collapsed")

# --- BEK KURUMSAL CSS TASARIMI ---
st.markdown("""
<style>
    /* Genel Arka Plan ve Metin Rengi */
    .stApp {
        background-color: #f4f6f9;
        color: #2c3e50;
    }
    /* Başlıklar - BEK Lacivert */
    h1, h2, h3, h4, h5, h6 {
        color: #00458b !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    /* Buton Tasarımları */
    .stButton>button {
        background-color: #00458b !important;
        color: white !important;
        border-radius: 6px;
        border: none;
        padding: 8px 16px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #003366 !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    /* İndirme Butonu Özelleştirmesi (Tasarruf vb. seçme butonları) */
    .stDownloadButton>button {
        background-color: #27ae60 !important;
        color: white !important;
    }
    .stDownloadButton>button:hover {
        background-color: #219653 !important;
    }
    /* Metrik Değerleri Rengi */
    div[data-testid="stMetricValue"] {
        color: #00458b;
    }
</style>
""", unsafe_allow_html=True)

# KPI'lar için format
def format_tr(val):
    if pd.isna(val) or val == 0:
        return "0 ₺"
    s = f"{val:,.0f}"
    return s.replace(",", ".") + " ₺"

def table_format_tr(val):
    if pd.isna(val): return ""
    try:
        s = f"{val:,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return val

options_ui = [50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 99]

if "hedef_butce" not in st.session_state:
    st.session_state["hedef_butce"] = 6200000000

# --- 1. ŞİFRE EKRANI ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.title("🏢 BEK Tedarik Zinciri - Güvenli Giriş")
        st.markdown("Sisteme erişmek için lütfen yönetici şifrenizi giriniz.")
        password = st.text_input("Şifre", type="password")
        if st.button("Sisteme Giriş Yap"):
            if password == "Eren12345":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("Hatalı Şifre! Lütfen tekrar deneyiniz.")
        return False
    return True

# --- EXCEL OLUŞTURMA FONKSİYONU (OPENPYXL İLE HATASIZ) ---
def create_formatted_excel(df_export, cols_to_format):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Senaryo_Sonucu')
        worksheet = writer.sheets['Senaryo_Sonucu']
        
        # openpyxl ile sütun genişliği ve sayı formatı ayarı
        for col_idx, col_name in enumerate(df_export.columns, 1): # Excel sütunları 1'den başlar
            if col_name in cols_to_format:
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 20
                for row_idx in range(2, len(df_export) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.number_format = '#,##0.00 "₺"'
            else:
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15
                
    return buffer.getvalue()

if check_password():
    st.title("📊 BEK Envanter & Bütçe Optimizasyon Paneli")
    st.markdown("Stok parametrelerinizi analiz edin, senaryoları inceleyin ve dilediğiniz bütçe planını tek tıkla raporlayın.")
    
    # --- 2. DOSYA YÜKLEME ---
    st.markdown("### 📑 1. Veri Yükleme")
    uploaded_files = st.file_uploader(
        "'Emniyet Seviyesi' Excel dosyasını seçiniz", 
        type=["xlsx", "xls", "xlsb"], accept_multiple_files=True
    )
    
    if uploaded_files:
        uploaded_file = uploaded_files[-1]
        try:
            engine_choice = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else None
            df = pd.read_excel(uploaded_file, sheet_name='Stok seviyesi', skiprows=6, engine=engine_choice)
            
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(subset=['Ürün Kodu']).copy()
            
            numeric_cols = ['Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            pareto_col_name = 'Defa Pareto' if 'Defa Pareto' in df.columns else ('Defa ABC' if 'Defa ABC' in df.columns else None)
            if pareto_col_name:
                pareto_series = df[pareto_col_name].fillna('BOŞ').astype(str).str.strip().str.upper()
                pareto_series = pareto_series.replace(['NAN', 'NONE', ''], 'BOŞ')
            else:
                pareto_series = pd.Series(['BOŞ'] * len(df), index=df.index)
            df['Temiz_Pareto'] = pareto_series

            old_levels_all = df['Emniyet Seviyesi'].clip(50, 99) / 100.0
            z_old_all = old_levels_all.apply(lambda p: NormalDist().inv_cdf(p))
            z_old_all = np.where(z_old_all == 0, 1e-5, z_old_all)
            df['Hedef_Min_Fark'] = df['Hedef TL'] - df['Min TL']
            
            st.success(f"Sistem Bilgisi: '{uploaded_file.name}' başarıyla işlendi.")
            
            # ---------------------------------------------------------
            # YÖNETİCİ ÖZETİ - GÜNCEL DURUM
            # ---------------------------------------------------------
            st.markdown("---")
            st.markdown("### 📈 Yönetici Özeti: Mevcut Envanter Durumu")
            
            mevcut_optimum_toplam = df['Optimum TL'].sum()
            mevcut_fazla_toplam = df['Fazla TL'].sum()
            mevcut_stok_toplam = df['Stok TL'].sum()
            
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
            st.markdown("### ⚙️ 2. Bütçeye Göre Stratejik Pareto Planlama")
            st.markdown("Analiz edilen senaryolardan size uygun olanını altındaki butona tıklayarak sonuçları doğrudan raporlayabilirsiniz.")
            
            col_rec1, col_rec2 = st.columns([2, 1])
            with col_rec1:
                analiz_butce = st.number_input(
                    "Hedeflediğiniz Bütçe Sınırını Giriniz (TL):",
                    value=int(st.session_state["hedef_butce"]),
                    step=50000000, format="%d"
                )
            with col_rec2:
                st.write("")
                st.write("")
                run_analysis = st.button("Analizi Başlat (8 Kademeli Senaryo)", use_container_width=True)

            if run_analysis:
                st.session_state["hedef_butce"] = analiz_butce
                
                with st.spinner("Kurumsal senaryolar hazırlanıyor..."):
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

                    p_indices = ["1. Tasarruf (Min)", "2. Ekonomik Alt", "3. Ekonomik Üst", "4. Yüksek Hizmet", "5. Bütçe Sınırı", "6. Bütçe Aşımı (%2)", "7. Bütçe Aşımı (%5)", "8. Max Fırsat (%10)"]
                    money_cols = ['Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL', 'Yeni Min TL', 'Yeni Optimum TL', 'Yeni Hedef TL', 'Yeni Fazla TL', 'Yeni Hedef Değer']
                    
                    st.markdown("#### 📑 Analiz Edilen Stratejik Senaryolar")
                    
                    # İLK SATIR (1-4)
                    row1 = st.columns(4)
                    for idx_i in range(4):
                        if idx_i >= len(selected_scenarios): continue
                        scen = selected_scenarios[idx_i]
                        
                        df_export = df[['Ürün Kodu', 'Ürün', pareto_col_name, 'Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL']].copy()
                        level_map_scen = {'A': scen['A'], 'B': scen['B'], 'C': scen['C'], 'D': scen['D'], 'E': scen['E'], 'BOŞ': scen['BOS']}
                        
                        df_export['Yeni_Emniyet_Seviyesi'] = df['Temiz_Pareto'].map(level_map_scen).fillna(scen['BOS'])
                        z_new_scen = (df_export['Yeni_Emniyet_Seviyesi']/100.0).apply(lambda p: NormalDist().inv_cdf(p))
                        
                        df_export['Yeni Min TL'] = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] * (z_new_scen / z_old_all))
                        df_export['Yeni Hedef TL'] = df_export['Yeni Min TL'] + df['Hedef_Min_Fark']
                        df_export['Yeni Optimum TL'] = (df_export['Yeni Min TL'] + df_export['Yeni Hedef TL']) / 2.0
                        df_export['Yeni Fazla TL'] = np.maximum(0, df['Stok TL'] - df_export['Yeni Hedef TL'])
                        df_export['Yeni Hedef Değer'] = df_export['Yeni Optimum TL'] + df_export['Yeni Fazla TL']
                        
                        scen_guncel_hedef = df_export['Yeni Hedef Değer'].sum()
                        excel_data = create_formatted_excel(df_export, money_cols)
                        
                        with row1[idx_i]:
                            st.subheader(p_indices[idx_i])
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            st.metric("Hedef Değer", format_tr(scen_guncel_hedef))
                            
                            st.download_button(
                                label="📥 Raporu İndir (Excel)",
                                data=excel_data,
                                file_name=f"BEK_Senaryo_{idx_i+1}_{int(scen_guncel_hedef)}_TL.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{idx_i}"
                            )

                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # İKİNCİ SATIR (5-8)
                    row2 = st.columns(4)
                    for idx_i in range(4, 8):
                        if idx_i >= len(selected_scenarios): continue
                        scen = selected_scenarios[idx_i]
                        
                        df_export = df[['Ürün Kodu', 'Ürün', pareto_col_name, 'Emniyet Seviyesi', 'Min TL', 'Optimum TL', 'Hedef TL', 'Stok TL', 'Fazla TL']].copy()
                        level_map_scen = {'A': scen['A'], 'B': scen['B'], 'C': scen['C'], 'D': scen['D'], 'E': scen['E'], 'BOŞ': scen['BOS']}
                        
                        df_export['Yeni_Emniyet_Seviyesi'] = df['Temiz_Pareto'].map(level_map_scen).fillna(scen['BOS'])
                        z_new_scen = (df_export['Yeni_Emniyet_Seviyesi']/100.0).apply(lambda p: NormalDist().inv_cdf(p))
                        
                        df_export['Yeni Min TL'] = np.where(df['Emniyet Seviyesi'] <= 50, 0, df['Min TL'] * (z_new_scen / z_old_all))
                        df_export['Yeni Hedef TL'] = df_export['Yeni Min TL'] + df['Hedef_Min_Fark']
                        df_export['Yeni Optimum TL'] = (df_export['Yeni Min TL'] + df_export['Yeni Hedef TL']) / 2.0
                        df_export['Yeni Fazla TL'] = np.maximum(0, df['Stok TL'] - df_export['Yeni Hedef TL'])
                        df_export['Yeni Hedef Değer'] = df_export['Yeni Optimum TL'] + df_export['Yeni Fazla TL']
                        
                        scen_guncel_hedef = df_export['Yeni Hedef Değer'].sum()
                        excel_data = create_formatted_excel(df_export, money_cols)
                        
                        with row2[idx_i - 4]:
                            st.subheader(p_indices[idx_i])
                            st.markdown(f"**A:** %{scen['A']} | **B:** %{scen['B']} | **C:** %{scen['C']}<br>**D:** %{scen['D']} | **E:** %{scen['E']} | **Boş:** %{scen['BOS']}", unsafe_allow_html=True)
                            st.markdown("---")
                            
                            delta_val = float(scen_guncel_hedef) - float(analiz_butce)
                            st.metric("Hedef Değer", format_tr(scen_guncel_hedef), delta=f"{format_tr(abs(delta_val))} Fark", delta_color="inverse" if delta_val > 0 else "normal")
                            
                            st.download_button(
                                label="📥 Raporu İndir (Excel)",
                                data=excel_data,
                                file_name=f"BEK_Senaryo_{idx_i+1}_{int(scen_guncel_hedef)}_TL.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{idx_i}"
                            )

        except Exception as e:
            st.error(f"Sistem Hatası: {e}")
