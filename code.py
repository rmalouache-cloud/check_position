import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule

st.set_page_config(page_title="CKD Position Validator", layout="wide")

# Style CSS personnalisé pour Streamlit
st.markdown("""
<style>
    /* Style pour le tableau Streamlit */
    .dataframe {
        font-size: 14px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    /* Animation pour les cartes de statistiques */
    .stMetric {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("# 🎨 Vérificateur de Positions CKD")
st.markdown("Vérifie que le nombre de positions correspond à la quantité pour les composants CKD uniquement")

# Initialiser session state
if 'show_problems' not in st.session_state:
    st.session_state.show_problems = False
if 'results_df' not in st.session_state:
    st.session_state.results_df = None
if 'verification_done' not in st.session_state:
    st.session_state.verification_done = False
if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None

# Upload du fichier BOM
old_file = st.file_uploader("📂 Upload votre fichier BOM", type=["xlsx"])

# Vérifier si un nouveau fichier a été uploadé
if old_file is not None:
    if st.session_state.current_file_name != old_file.name:
        # Nouveau fichier - réinitialiser les résultats
        st.session_state.results_df = None
        st.session_state.verification_done = False
        st.session_state.show_problems = False
        st.session_state.current_file_name = old_file.name

def extract_ckd_components(df):
    """Extrait uniquement les composants CKD"""
    start_idx = None
    for idx, desc in enumerate(df['Description']):
        if 'ASS\'Y - MAIN BOARD（CKD）' in str(desc).upper() or 'ASSY - MAIN BOARD（CKD）' in str(desc).upper():
            start_idx = idx
            break
    
    end_idx = None
    for idx, desc in enumerate(df['Description']):
        if 'BARCODE LABEL' in str(desc).upper():
            end_idx = idx
            break
    
    if start_idx is not None and end_idx is not None:
        return df.iloc[start_idx:end_idx+1].copy()
    elif start_idx is not None:
        return df.iloc[start_idx:].copy()
    else:
        return pd.DataFrame()

def safe_join(x):
    if not isinstance(x, list):
        return str(x) if pd.notna(x) else ""
    return ", ".join(str(i) for i in x if pd.notna(i))

def extract_positions(bom_text):
    positions = []
    if bom_text and str(bom_text) != "nan":
        bom_text_str = str(bom_text)
        raw_positions = bom_text_str.split(',')
        for pos in raw_positions:
            pos = pos.strip()
            pos = pos.replace('[', '').replace(']', '').replace("'", "").replace('"', "").strip()
            if pos and pos != "nan":
                positions.append(pos)
    return positions

def is_non_component(description):
    non_components = [
        'ASS\'Y - MAIN BOARD（CKD）',
        'ASSY - MAIN BOARD（CKD）',
        'ASS\'Y - MAIN BOARD',
        'ASSY - MAIN BOARD',
        'ASS\'Y - MAIN BOARD（SMT）',
        'ASSY - MAIN BOARD（SMT）',
        'PCB',
        'THERMAL CONDUCTIVE',
        'BARCODE LABEL'
    ]
    desc_upper = str(description).upper()
    for non_comp in non_components:
        if non_comp.upper() in desc_upper:
            return True
    return False

def validate_ckd_positions(df):
    results = []
    
    for idx, row in df.iterrows():
        pn = row.get("PN", "")
        description = row.get("Description", "")
        bom_text = row.get("BOM text", "")
        qty = row.get("bom_qty", 0)
        
        is_non_comp = is_non_component(description)
        
        try:
            if isinstance(qty, str):
                qty = qty.replace(",", ".")
            qty = float(qty) if qty else 0
        except:
            qty = 0
        
        positions = extract_positions(bom_text)
        nb_positions = len(positions)
        positions_str = safe_join(positions)
        
        # Déterminer le résultat
        if is_non_comp:
            result_detail = "📌 NO NEED / NOT APPLICABLE"
            result_class = "no-need"
        elif nb_positions == 0 and qty == 0:
            result_detail = "✅ VIDE"
            result_class = "vide"
        elif nb_positions == 0 and qty > 0:
            result_detail = "❌ ERREUR - Aucune position"
            result_class = "erreur"
        elif nb_positions == qty:
            result_detail = "✅ CONFORME"
            result_class = "conforme"
        elif nb_positions < qty:
            result_detail = f"⚠️ MANQUE - {int(qty - nb_positions)} position(s) manquante(s)"
            result_class = "manque"
        else:
            result_detail = f"⚠️ TROP - {int(nb_positions - qty)} position(s) en excès"
            result_class = "trop"
        
        results.append({
            "PN": pn,
            "Description": description,
            "QTY": int(qty) if qty == int(qty) else qty,
            "Position": positions_str if positions_str else "—",
            "QTY Calculated": nb_positions,
            "Result": result_detail,
            "Result_Class": result_class
        })
    
    return pd.DataFrame(results)

def export_to_colored_excel(df, filename):
    """Exporte vers Excel avec des couleurs vives et professionnelles"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Verification_CKD', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Verification_CKD']
        
        # Appliquer les couleurs à chaque ligne selon son statut
        for row_idx in range(2, worksheet.max_row + 1):
            result_cell = worksheet.cell(row=row_idx, column=6)
            result_text = str(result_cell.value)
            
            # Déterminer la couleur en fonction du texte
            if "CONFORME" in result_text and "NO NEED" not in result_text:
                color = "92D050"
                font_color = "FFFFFF"
            elif "ERREUR" in result_text:
                color = "FF6B6B"
                font_color = "FFFFFF"
            elif "MANQUE" in result_text:
                color = "FFB347"
                font_color = "FFFFFF"
            elif "TROP" in result_text:
                color = "FFB347"
                font_color = "FFFFFF"
            elif "NO NEED" in result_text:
                color = "5B9BD5"
                font_color = "FFFFFF"
            elif "VIDE" in result_text:
                color = "C5E0B4"
                font_color = "000000"
            else:
                color = "FFFFFF"
                font_color = "000000"
            
            # Appliquer le remplissage et la couleur de police
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            font = Font(color=font_color, bold=True)
            
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_idx, column=col)
                cell.fill = fill
                if font_color == "FFFFFF":
                    cell.font = font
            
            # Ajouter des bordures
            thin_border = Border(
                left=Side(style='thin', color='DDDDDD'),
                right=Side(style='thin', color='DDDDDD'),
                top=Side(style='thin', color='DDDDDD'),
                bottom=Side(style='thin', color='DDDDDD')
            )
            for col in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row_idx, column=col).border = thin_border
        
        # Style de l'en-tête
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        
        for col in range(1, worksheet.max_column + 1):
            header_cell = worksheet.cell(row=1, column=col)
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Ajuster les largeurs des colonnes
        column_widths = {
            'A': 20,
            'B': 35,
            'C': 10,
            'D': 30,
            'E': 15,
            'F': 40,
        }
        
        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width
        
        worksheet.freeze_panes = 'A2'
    
    output.seek(0)
    return output

def color_result_css(val):
    """Style CSS pour le tableau Streamlit - plus élégant"""
    if "CONFORME" in str(val) and "NO NEED" not in str(val):
        return 'background: linear-gradient(135deg, #92D050 0%, #7CB83D 100%); color: white; font-weight: bold; border-radius: 5px; padding: 5px;'
    elif "ERREUR" in str(val):
        return 'background: linear-gradient(135deg, #FF6B6B 0%, #E55555 100%); color: white; font-weight: bold; border-radius: 5px; padding: 5px;'
    elif "MANQUE" in str(val):
        return 'background: linear-gradient(135deg, #FFB347 0%, #FF9500 100%); color: white; font-weight: bold; border-radius: 5px; padding: 5px;'
    elif "TROP" in str(val):
        return 'background: linear-gradient(135deg, #FFB347 0%, #FF9500 100%); color: white; font-weight: bold; border-radius: 5px; padding: 5px;'
    elif "NO NEED" in str(val):
        return 'background: linear-gradient(135deg, #5B9BD5 0%, #3A7BC8 100%); color: white; font-weight: bold; border-radius: 5px; padding: 5px;'
    elif "VIDE" in str(val):
        return 'background: linear-gradient(135deg, #C5E0B4 0%, #A8D08D 100%); color: #2C3E50; font-weight: bold; border-radius: 5px; padding: 5px;'
    return ''

if old_file:
    try:
        df = pd.read_excel(old_file)
        df.columns = df.columns.str.strip()
        
        with st.expander("📋 Aperçu du fichier chargé"):
            st.dataframe(df.head(10), use_container_width=True)
        
        required_cols = ["bom_qty", "BOM text"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Colonnes manquantes: {', '.join(missing_cols)}")
        else:
            with st.spinner("Extraction des composants CKD..."):
                ckd_df = extract_ckd_components(df)
            
            if ckd_df.empty:
                st.warning("⚠️ Aucun composant CKD trouvé dans ce fichier")
                # Afficher un aperçu pour aider l'utilisateur
                st.info("Voici les 20 premières descriptions du fichier:")
                st.dataframe(df[["Description"]].head(20), use_container_width=True)
            else:
                st.success(f"✅ {len(ckd_df)} composants CKD extraits")
                
                if st.button("🔍 Vérifier les positions CKD", use_container_width=True):
                    with st.spinner("Vérification en cours..."):
                        st.session_state.results_df = validate_ckd_positions(ckd_df)
                        st.session_state.verification_done = True
                        st.session_state.show_problems = False
                        st.rerun()
                
                # Afficher les résultats si la vérification a été faite
                if st.session_state.verification_done and st.session_state.results_df is not None:
                    results_df = st.session_state.results_df
                    
                    # Statistiques avec style
                    st.subheader("📊 Résumé de la vérification CKD")
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    total = len(results_df)
                    conformes = len(results_df[results_df["Result"].str.contains("CONFORME", na=False) & ~results_df["Result"].str.contains("NO NEED", na=False)])
                    erreurs = len(results_df[results_df["Result"].str.contains("ERREUR", na=False)])
                    manques = len(results_df[results_df["Result"].str.contains("MANQUE", na=False)])
                    trop = len(results_df[results_df["Result"].str.contains("TROP", na=False)])
                    
                    with col1:
                        st.metric("📊 Total", total)
                    with col2:
                        st.metric("✅ Conformes", conformes)
                    with col3:
                        st.metric("❌ Erreurs", erreurs, delta="À corriger" if erreurs > 0 else None)
                    with col4:
                        st.metric("⚠️ Manque", manques)
                    with col5:
                        st.metric("⚠️ Trop", trop)
                    
                    st.markdown("---")
                    
                    # Checkbox pour filtrer
                    def toggle_filter():
                        st.session_state.show_problems = not st.session_state.show_problems
                        st.rerun()
                    
                    col_check, col_info = st.columns([1, 3])
                    with col_check:
                        st.checkbox("⚠️ Afficher uniquement les lignes non conformes", 
                                   value=st.session_state.show_problems,
                                   on_change=toggle_filter,
                                   key="filter_checkbox")
                    
                    st.subheader("🔍 Détail de la vérification CKD")
                    
                    # Sélection des colonnes
                    display_cols = ["PN", "Description", "QTY", "Position", "QTY Calculated", "Result"]
                    
                    # Filtrer ou non
                    if st.session_state.show_problems:
                        filtered_df = results_df[
                            (results_df["Result"].str.contains("ERREUR", na=False)) |
                            (results_df["Result"].str.contains("MANQUE", na=False)) |
                            (results_df["Result"].str.contains("TROP", na=False))
                        ][display_cols].copy()
                        
                        if len(filtered_df) > 0:
                            st.warning(f"⚠️ {len(filtered_df)} ligne(s) non conforme(s) sur {len(results_df)} total")
                            styled_filtered = filtered_df.style.map(color_result_css, subset=['Result'])
                            st.dataframe(styled_filtered, use_container_width=True)
                        else:
                            st.success("✅ Aucune ligne non conforme trouvée !")
                            st.balloons()  # Animation de succès
                            st.info(f"✨ Toutes les {len(results_df)} lignes sont conformes ou non applicables")
                    else:
                        styled_df = results_df[display_cols].copy().style.map(color_result_css, subset=['Result'])
                        st.dataframe(styled_df, use_container_width=True)
                    
                    # Export Excel
                    colored_excel = export_to_colored_excel(results_df[display_cols], "verification_positions_CKD.xlsx")
                    
                    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                    with col_btn2:
                        st.download_button(
                            "📥 Télécharger le rapport Excel (coloré)",
                            colored_excel,
                            f"verification_CKD_{st.session_state.current_file_name.replace('.xlsx', '')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
    except Exception as e:
        st.error(f"❌ Erreur: {str(e)}")
else:
    st.info("👆 Veuillez uploader un fichier Excel")
    # Réinitialiser quand aucun fichier n'est chargé
    st.session_state.results_df = None
    st.session_state.verification_done = False
    st.session_state.current_file_name = None
