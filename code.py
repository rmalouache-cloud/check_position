import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font, Border, Side, Alignment

st.set_page_config(page_title="CKD Position Validator", layout="wide")
st.markdown("# 📺 Vérificateur de Positions CKD")
st.markdown("Vérifie que le nombre de positions correspond à la quantité pour les composants CKD uniquement")

# Upload du fichier BOM
old_file = st.file_uploader("📂 Upload votre fichier BOM", type=["xlsx"])

def extract_ckd_components(df):
    """
    Extrait uniquement les composants CKD
    CKD: de "ASS'Y - MAIN BOARD（CKD）" jusqu'à "BARCODE LABEL"
    """
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
    """Convertit une liste en chaîne de caractères"""
    if not isinstance(x, list):
        return str(x) if pd.notna(x) else ""
    return ", ".join(str(i) for i in x if pd.notna(i))

def extract_positions(bom_text):
    """
    Extrait les positions individuelles depuis BOM text
    """
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
    """
    Vérifie si la ligne n'est pas un composant (maison mère, PCB, etc.)
    """
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
    """
    Vérifie la cohérence entre les positions et la quantité pour CKD
    """
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
        
        # Déterminer le résultat et le statut pour le filtrage
        if is_non_comp:
            result = "📌 NO NEED / NOT APPLICABLE"
            status_category = "non_applicable"
        elif nb_positions == 0 and qty == 0:
            result = "✅ VIDE"
            status_category = "conforme"
        elif nb_positions == 0 and qty > 0:
            result = "❌ ERREUR - Aucune position"
            status_category = "non_conforme"
        elif nb_positions == qty:
            result = "✅ CONFORME"
            status_category = "conforme"
        elif nb_positions < qty:
            result = f"⚠️ MANQUE - {int(qty - nb_positions)} position(s) manquante(s)"
            status_category = "non_conforme"
        else:
            result = f"⚠️ TROP - {int(nb_positions - qty)} position(s) en excès"
            status_category = "non_conforme"
        
        results.append({
            "PN": pn,
            "Description": description,
            "QTY": int(qty) if qty == int(qty) else qty,
            "Position": positions_str,
            "QTY Calculated": nb_positions,
            "Result": result,
            "Status_Category": status_category
        })
    
    return pd.DataFrame(results)

def export_to_colored_excel(df, filename):
    """
    Exporte le DataFrame vers Excel avec couleurs
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Verification_CKD', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Verification_CKD']
        
        color_map = {
            "📌 NO NEED / NOT APPLICABLE": "D9E1F2",
            "✅ CONFORME": "C6EFCE",
            "❌ ERREUR - Aucune position": "FFC7CE",
            "✅ VIDE": "E2EFDA",
        }
        warning_color = "FFEB9C"
        
        for row in range(2, worksheet.max_row + 1):
            status_cell = worksheet.cell(row=row, column=6)
            status = status_cell.value
            
            if status in color_map:
                color = color_map[status]
            elif "MANQUE" in str(status) or "TROP" in str(status):
                color = warning_color
            else:
                color = "FFFFFF"
            
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for col in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row, column=col).fill = fill
            
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            for col in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row, column=col).border = thin_border
        
        for col in range(1, worksheet.max_column + 1):
            header_cell = worksheet.cell(row=1, column=col)
            header_cell.font = Font(bold=True)
            header_cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_cell.font = Font(bold=True, color="FFFFFF")
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output

if old_file:
    try:
        df = pd.read_excel(old_file)
        df.columns = df.columns.str.strip()
        
        with st.expander("📋 Aperçu du fichier chargé"):
            st.dataframe(df.head(10), use_container_width=True)
        
        required_cols = ["bom_qty", "BOM text"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Colonnes manquantes dans le fichier: {', '.join(missing_cols)}")
            st.info(f"Les colonnes disponibles sont: {', '.join(df.columns)}")
        else:
            with st.spinner("Extraction des composants CKD..."):
                ckd_df = extract_ckd_components(df)
            
            if ckd_df.empty:
                st.warning("⚠️ Aucun composant CKD trouvé dans le fichier")
                st.info("Assurez-vous que votre fichier contient la ligne 'ASS'Y - MAIN BOARD（CKD）' et 'BARCODE LABEL'")
            else:
                st.success(f"✅ {len(ckd_df)} composants CKD extraits")
                
                if st.button("🔍 Vérifier les positions CKD", use_container_width=True):
                    with st.spinner("Vérification en cours..."):
                        results_df = validate_ckd_positions(ckd_df)
                        
                        # Statistiques
                        st.subheader("📊 Résumé de la vérification CKD")
                        
                        col1, col2, col3, col4, col5 = st.columns(5)
                        total = len(results_df)
                        conformes = len(results_df[results_df["Result"] == "✅ CONFORME"])
                        erreurs = len(results_df[results_df["Result"] == "❌ ERREUR - Aucune position"])
                        manques = len(results_df[results_df["Result"].str.contains("MANQUE", na=False)])
                        trop = len(results_df[results_df["Result"].str.contains("TROP", na=False)])
                        no_need = len(results_df[results_df["Result"] == "📌 NO NEED / NOT APPLICABLE"])
                        
                        with col1:
                            st.metric("📊 Total", total)
                        with col2:
                            st.metric("✅ Conformes", conformes)
                        with col3:
                            st.metric("❌ Erreurs", erreurs)
                        with col4:
                            st.metric("⚠️ Manque", manques)
                        with col5:
                            st.metric("⚠️ Trop", trop)
                        
                        st.markdown("---")
                        
                        # Sélectionner les colonnes à afficher
                        display_cols = ["PN", "Description", "QTY", "Position", "QTY Calculated", "Result"]
                        display_df = results_df[display_cols].copy()
                        
                        # Option pour filtrer les problèmes uniquement
                        show_problems_only = st.checkbox("⚠️ Afficher uniquement les lignes non conformes", value=False)
                        
                        st.subheader("🔍 Détail de la vérification CKD")
                        
                        # CORRECTION DU FILTRE
                        if show_problems_only:
                            # Filtrer les lignes non conformes (ERREUR, MANQUE, TROP)
                            mask = (
                                (display_df["Result"].str.contains("ERREUR", na=False)) |
                                (display_df["Result"].str.contains("MANQUE", na=False)) |
                                (display_df["Result"].str.contains("TROP", na=False))
                            )
                            filtered_df = display_df[mask].copy()
                            
                            if len(filtered_df) > 0:
                                st.warning(f"⚠️ {len(filtered_df)} ligne(s) non conforme(s) sur {len(display_df)} total")
                                
                                # Appliquer le style
                                def color_result(val):
                                    if "ERREUR" in str(val):
                                        return 'background-color: #FFC7CE'
                                    elif "MANQUE" in str(val) or "TROP" in str(val):
                                        return 'background-color: #FFEB9C'
                                    return ''
                                
                                styled_filtered_df = filtered_df.style.map(color_result, subset=['Result'])
                                st.dataframe(styled_filtered_df, use_container_width=True)
                            else:
                                st.success("✅ Aucune ligne non conforme trouvée !")
                                st.info(f"Toutes les {len(display_df)} lignes sont conformes ou non applicables")
                        else:
                            # Afficher toutes les lignes avec style
                            def color_result_all(val):
                                if "CONFORME" in str(val):
                                    return 'background-color: #C6EFCE'
                                elif "ERREUR" in str(val):
                                    return 'background-color: #FFC7CE'
                                elif "MANQUE" in str(val) or "TROP" in str(val):
                                    return 'background-color: #FFEB9C'
                                elif "NO NEED" in str(val):
                                    return 'background-color: #D9E1F2'
                                elif "VIDE" in str(val):
                                    return 'background-color: #E2EFDA'
                                return ''
                            
                            styled_df = display_df.style.map(color_result_all, subset=['Result'])
                            st.dataframe(styled_df, use_container_width=True)
                        
                        # Export Excel
                        colored_excel = export_to_colored_excel(display_df, "verification_positions_CKD.xlsx")
                        
                        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                        with col_btn2:
                            st.download_button(
                                "📥 Télécharger le rapport Excel (coloré)",
                                colored_excel,
                                "verification_positions_CKD.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                        with st.expander("ℹ️ Comprendre la vérification CKD"):
                            st.markdown("""
                            ### Comment fonctionne la vérification ?
                            
                            La vérification compare le **nombre de positions** avec la **quantité**.
                            
                            | Statut | Signification |
                            |--------|---------------|
                            | ✅ CONFORME | Nombre de positions = Quantité |
                            | ❌ ERREUR | Quantité > 0 mais aucune position |
                            | ⚠️ MANQUE | Nombre de positions < Quantité |
                            | ⚠️ TROP | Nombre de positions > Quantité |
                            | 📌 NO NEED | Élément non applicable (PCB, etc.) |
                            | ✅ VIDE | Quantité = 0 et pas de positions |
                            """)
                        
    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture du fichier: {str(e)}")
else:
    st.info("👆 Veuillez uploader un fichier Excel pour commencer la vérification")
    
    with st.expander("📖 Format attendu du fichier"):
        st.markdown("""
        ### Colonnes requises:
        - `PN` : Référence du composant
        - `Description` : Description
        - `bom_qty` : Quantité
        - `BOM text` : Positions (séparées par des virgules)
        
        ### Structure CKD:
        - Début: `ASS'Y - MAIN BOARD（CKD）`
        - Fin: `BARCODE LABEL`
        """)
