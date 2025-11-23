import streamlit as st
import pandas as pd
import io

st.title("Traitement de fichier Excel")
st.write("Uploadez votre fichier Excel pour obtenir les fichiers traités")

uploaded_file = st.file_uploader("Choisir un fichier Excel", type=['xlsx'])

last_date = st.date_input(
    "Dernière date de traitement (optionnel)",
    value=None,
    help="Si renseignée, seules les lignes avec une date supérieure seront traitées"
)

if uploaded_file is not None:
    if st.button("Traiter le fichier"):
        with st.spinner("Traitement en cours..."):
            # Import the Excel file
            data = pd.read_excel(uploaded_file, sheet_name='Sheet1', skiprows=3, dtype='string')
            
            # Drop the first unnamed column if it exists
            if 'Unnamed: 0' in data.columns:
                data = data.drop(columns=['Unnamed: 0'])
            
            # Filter by date if last_date is provided
            if last_date is not None:
                data['Date_parsed'] = pd.to_datetime(data['Date'], format='%d/%m/%y', errors='coerce')
                last_date_pd = pd.Timestamp(last_date)
                data = data[data['Date_parsed'] > last_date_pd].copy()
                data = data.drop(columns=['Date_parsed'])
                st.info(f"Filtrage appliqué: {len(data)} lignes après la date {last_date.strftime('%d/%m/%Y')}")
            
            # Add group index based on client headers
            is_date_nan = data['Date'].isna()
            is_header_client = is_date_nan & data['CJ Fol'].notna()
            data['group_id'] = is_header_client.cumsum()
            
            # Filter to keep only specific columns
            data = data[['group_id', 'Date', 'CJ Fol', 'Pièce', 'Débit', 'Lt.', 'Crédit']]
            
            # Normalize CJ Fol names
            data['CJ Fol'] = data['CJ Fol'].replace({'BQ  010': 'BQ', 'EAR 000': 'BQ', 'BQ  000': 'BQ', 'OD  000': 'BQ'})
            
            # Convert Débit and Crédit to float
            data['Débit'] = pd.to_numeric(data['Débit'], errors='coerce').fillna(0)
            data['Crédit'] = pd.to_numeric(data['Crédit'], errors='coerce').fillna(0)
            data['balance'] = data['Débit'] - data['Crédit']
            
            # Add date_reglement column
            data['date_reglement'] = None
            
            # Group by group_id and Lt. to process efficiently
            grouped = data[data['Lt.'].notna()].groupby(['group_id', 'Lt.'], sort=False)
            
            for (group_id, lt_val), group in grouped:
                indices = group.index.tolist()
                n = len(indices)
                
                i = 0
                while i < n:
                    cumsum = 0
                    block_start = i
                    
                    for j in range(i, n):
                        cumsum += data.loc[indices[j], 'balance']
                        
                        if abs(cumsum) < 0.01:
                            block_indices = indices[block_start:j+1]
                            block_df = data.loc[block_indices]
                            
                            bq_rows = block_df[block_df['CJ Fol'] == 'BQ']
                            vt_rows = block_df[
                                (block_df['CJ Fol'] == 'VT  000') & 
                                (block_df['Pièce'].notna()) & 
                                (block_df['Pièce'].str.startswith('A', na=False))
                            ]
                            
                            piece_indices = block_df[block_df['Pièce'].notna()].index
                            
                            if len(piece_indices) > 0:
                                if len(bq_rows) > 0:
                                    data.loc[piece_indices, 'date_reglement'] = bq_rows.iloc[0]['Date']
                                elif len(vt_rows) > 0:
                                    data.loc[piece_indices, 'date_reglement'] = vt_rows.iloc[0]['Date']
                            
                            i = j + 1
                            break
                    else:
                        i += 1
            
            # Stats
            total_with_date = data['date_reglement'].notna().sum()
            total_pieces = len(data[data['Pièce'].notna()])
            
            st.success(f"✓ Traitement terminé! {total_with_date}/{total_pieces} pièces traitées")
            
            # Prepare Excel output
            excel_buffer = io.BytesIO()
            data.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            
            # Prepare TXT output
            final_df = data[data['date_reglement'].notna()].copy()
            final_df['date_formatted'] = final_df['date_reglement'].str.replace('/', '', regex=False)
            final_df['montant'] = final_df['Débit'] - final_df['Crédit']
            output_df = final_df[['Pièce', 'date_formatted', 'montant']]
            
            txt_buffer = io.StringIO()
            output_df.to_csv(txt_buffer, sep=' ', header=False, index=False)
            txt_content = txt_buffer.getvalue()
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Télécharger Excel",
                    data=excel_buffer,
                    file_name="resultat.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="📥 Télécharger TXT",
                    data=txt_content,
                    file_name="resultat.txt",
                    mime="text/plain"
                )
