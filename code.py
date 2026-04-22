import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font, Border, Side

st.set_page_config(page_title="CKD Position Validator", layout="wide")
st.markdown("# 📺 Vérificateur de Positions CKD")
st.markdown("Vérifie que le nombre de positions correspond à la quantité pour les composants CKD uniquement")

# Upload du fichier BOM
old_file = st.file_uploader("📂 Upload votre fichier BOM", type=["xlsx"])

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
        
        # Déterminer le résultat avec des codes simples
        if is_non_comp:
            result_text = "NO NEED"
            result_detail = "📌 NO NEED / NOT APPLICABLE"
            is_problem = False
        elif nb_positions == 0 and qty == 0:
            result_text = "VIDE"
            result_detail = "✅ VIDE"
            is_problem = False
        elif nb_positions == 0 and qty > 0:
            result_text = "ERREUR"
            result_detail = "❌ ERREUR - Aucune position"
            is_problem = True
        elif nb_positions == qty:
            result_text = "CONFORME"
            result_detail = "✅ CONFORME"
            is_problem = False
        elif nb_positions < qty:
            result_text = "MANQUE"
            result_detail = f"⚠️ MANQUE - {int(qty - nb_positions)} position(s) manquante(s)"
            is_problem = True
        else:
            result_text = "TROP"
            result_detail = f"⚠️ TROP - {int(nb_positions - qty)} position(s) en excès"
            is_problem = True
        
        results.append({
            "PN": pn,
            "Description": description,
            "QTY": int(qty) if qty == int(qty) else qty,
            "Position": positions_str,
            "QTY Calculated": nb_positions,
            "Result": result_detail,
            "Status": result_text,  # Colonne pour le filtrage
            "IsProblem": is_problem  # Colonne booléenne pour le filtrage
        })
    
    return pd.DataFrame(results)

def export_to_colored_excel(df, filename):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Enlever la colonne IsProblem et Status pour l'export
        export_df = df.drop(columns=['IsProblem', 'Status'], errors='ignore')
        export_df.to_excel(writer, sheet_name='Verification_CKD', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Verification_CKD']
        
        # Parcourir les lignes et colorer selon le résultat
        for row in range(2, worksheet.max_row + 1):
            status_cell = worksheet.cell(row=row, column=6)  # Colonne Result
            status = status_cell.value
            
            if "CONFORME" in str(status):
                color = "C6EFCE"  # Vert
            elif "ERREUR" in str(status):
                color = "FFC7CE"  # Rouge
            elif "MANQUE" in str(status) or "TROP" in str(status):
                color = "FFEB9C"  # Jaune
            elif "NO NEED" in str(status):
                color = "D9E1F2"  # Bleu
            elif "VIDE" in str(status):
                color = "E2EFDA"  # Vert clair
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
        
        # En-tête
        for col in range(1, worksheet.max_column + 1):
            header_cell = worksheet.cell(row=1, column=col)
            header_cell.font = Font(bold=True, color="FFFFFF")
            header_cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        
        # Ajuster les largeurs
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
            st.error(f"❌ Colonnes manquantes: {', '.join(missing_cols)}")
        else:
            with st.spinner("Extraction des composants CKD..."):
                ckd_df = extract_ckd_components(df)
            
            if ckd_df.empty:
                st.warning("⚠️ Aucun composant CKD trouvé")
            else:
                st.success(f"✅ {len(ckd_df)} composants CKD extraits")
                
                if st.button("🔍 Vérifier les positions CKD", use_container_width=True):
                    with st.spinner("Vérification en cours..."):
                        results_df = validate_ckd_positions(ckd_df)
                        
                        # Statistiques
                        st.subheader("📊 Résumé de la vérification CKD")
                        
                        col1, col2, col3, col4, col5 = st.columns(5)
                        total = len(results_df)
                        conformes = len(results_df[results_df["Status"] == "CONFORME"])
                        erreurs = len(results_df[results_df["Status"] == "ERREUR"])
                        manques = len(results_df[results_df["Status"] == "MANQUE"])
                        trop = len(results_df[results_df["Status"] == "TROP"])
                        no_need = len(results_df[results_df["Status"] == "NO NEED"])
                        
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
                        
                        # Sélection des colonnes à afficher
                        display_cols = ["PN", "Description", "QTY", "Position", "QTY Calculated", "Result"]
                        
                        # Checkbox pour filtrer
                        show_problems_only = st.checkbox("⚠️ Afficher uniquement les lignes non conformes", value=False)
                        
                        st.subheader("🔍 Détail de la vérification CKD")
                        
                        # FILTRAGE SIMPLE ET EFFICACE
                        if show_problems_only:
                            # Utiliser la colonne booléenne IsProblem
                            filtered_df = results_df[results_df["IsProblem"] == True][display_cols].copy()
                            
                            if len(filtered_df) > 0:
                                st.warning(f"⚠️ {len(filtered_df)} ligne(s) non conforme(s) sur {len(results_df)} total")
                                
                                # Style pour le tableau filtré
                                def color_problem(val):
                                    if "ERREUR" in str(val):
                                        return 'background-color: #FFC7CE'
                                    elif "MANQUE" in str(val) or "TROP" in str(val):
                                        return 'background-color: #FFEB9C'
                                    return ''
                                
                                st.dataframe(filtered_df.style.map(color_problem, subset=['Result']), use_container_width=True)
                            else:
                                st.success("✅ Aucune ligne non conforme trouvée !")
                        else:
                            # Afficher toutes les lignes
                            all_df = results_df[display_cols].copy()
                            
                            def color_all(val):
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
                            
                            st.dataframe(all_df.style.map(color_all, subset=['Result']), use_container_width=True)
                        
                        # Export Excel
                        colored_excel = export_to_colored_excel(results_df, "verification_positions_CKD.xlsx")
                        
                        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                        with col_btn2:
                            st.download_button(
                                "📥 Télécharger le rapport Excel (coloré)",
                                colored_excel,
                                "verification_positions_CKD.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
    except Exception as e:
        st.error(f"❌ Erreur: {str(e)}")
else:
    st.info("👆 Veuillez uploader un fichier Excel")
