import pandas as pd
import os
import glob
import numpy as np

# Config
BASE_DIR = r"C:\Users\marce\Documents\Atlantico Holding\Briefing_Automação"
SOURCE_DIR = os.path.join(BASE_DIR, "Planilha Base")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Helper to read single cells
def get_val(df, row, col):
    try:
        val = df.iloc[row, col]
        return val if pd.notna(val) else None
    except:
        return None

def clean_int(val):
    try:
        if pd.notna(val):
            return int(round(val))
        return None
    except:
        return None

def find_col(row, val, start_idx=0):
    for i in range(start_idx, len(row)):
        if str(row[i]) == val:
            return i
    return -1

def process_file(file_path):
    print(f"\nProcessing file: {file_path}")
    try:
        # Load 'Informações Principais'
        df_info = pd.read_excel(file_path, sheet_name='Informações Principais', header=None)

        # Extract Key Info
        campaign_id = get_val(df_info, 10, 2) 
        expert_name = get_val(df_info, 14, 2) 
        cpl_meta = get_val(df_info, 33, 2)    
        meta_vendas = get_val(df_info, 35, 2) 
        meta_faturamento = get_val(df_info, 36, 2) 

        account_meta = get_val(df_info, 20, 2)
        account_google = get_val(df_info, 21, 2)

        # Dates
        dates = {
            'Captação': (get_val(df_info, 41, 2), get_val(df_info, 42, 2)),
            'Aquecimento': (get_val(df_info, 46, 2), get_val(df_info, 47, 2)),
            'Lembrete': (get_val(df_info, 51, 2), get_val(df_info, 52, 2)),
            'Remarketing Vendas': (get_val(df_info, 66, 2), get_val(df_info, 67, 2)),
            'Distribuição de Conteúdo': (None, None),
            'Remarketing Aulas': (None, None),
            'Flash Opening': (None, None)
        }

        print(f"  Campaign ID: {campaign_id}")
        print(f"  Expert: {expert_name}")

        # --- File 1: [00.00][ID] Planejamento.xlsx ---
        file1_name = f"[00.00][{campaign_id}] Planejamento.xlsx"
        
        with pd.ExcelWriter(os.path.join(OUTPUT_DIR, file1_name)) as writer:
            # Sheet: Experts
            df_experts = pd.DataFrame({
                'Expert': ['Expert 1', 'Expert 2', 'Expert 3'],
                'Nome do Expert': [expert_name, get_val(df_info, 15, 2), get_val(df_info, 16, 2)]
            })
            df_experts.to_excel(writer, sheet_name='Experts', index=False)

            # Sheet: Datas das Fases
            fases_data = []
            for fase in ['Distribuição de Conteúdo', 'Captação', 'Aquecimento', 'Lembrete', 'Remarketing Aulas', 'Remarketing Vendas', 'Flash Opening']:
                start, end = dates.get(fase, (None, None))
                fases_data.append([fase, start, end])
            
            df_dates = pd.DataFrame(fases_data, columns=['Fase', 'Data de Início', 'Data de Fim'])
            df_dates.to_excel(writer, sheet_name='Datas das Fases', index=False)

            # Sheet: Meta de CPI (Empty)
            pd.DataFrame(columns=['Meta de CPI']).to_excel(writer, sheet_name='Meta de CPI', index=False)

            # Sheet: Meta de CPL
            pd.DataFrame({'Meta de CPL': [cpl_meta]}).to_excel(writer, sheet_name='Meta de CPL', index=False)


        # --- File 2: [01.00][ID] Campanha 1.xlsx ---
        file2_name = f"[01.00][{campaign_id}] Campanha 1.xlsx"
        
        with pd.ExcelWriter(os.path.join(OUTPUT_DIR, file2_name)) as writer:
            # Sheet: Metas de Venda
            pd.DataFrame({
                'Meta de Vendas': [meta_vendas],
                'Meta de Faturamento': [meta_faturamento]
            }).to_excel(writer, sheet_name='Metas de Venda', index=False)

            # Sheet: ID da Campanha
            pd.DataFrame({'ID da Campanha': [campaign_id]}).to_excel(writer, sheet_name='ID da Campanha', index=False)

            # Sheet: Ingresso
            pd.DataFrame(columns=['Ingresso']).to_excel(writer, sheet_name='Ingresso', index=False)


        # --- File 3: [01.01][ID][Campanha 1] Expert 1.xlsx ---
        file3_name = f"[01.01][{campaign_id}][Campanha 1] Expert 1.xlsx"
        
        # Load Source Data for File 3
        df_leads = pd.read_excel(file_path, sheet_name='Meta de Leads', header=None)
        
        leads_data = {
            'Tráfego': clean_int(get_val(df_leads, 4, 5)),
            'Meta Ads': clean_int(get_val(df_leads, 5, 5)),
            'YouTube Ads': clean_int(get_val(df_leads, 6, 5)),
            'Google Ads': clean_int(get_val(df_leads, 7, 5)),
            'Social': clean_int(get_val(df_leads, 8, 5)),
            'Mailing': clean_int(get_val(df_leads, 9, 5))
        }

        # Load Source Data for Investimento
        df_invest = pd.read_excel(file_path, sheet_name='Investimento', header=None)
        
        invest_map = {}
        phases = ['Distribuição de Conteúdo', 'Captação', 'Aquecimento', 'Lembrete', 'Remarketing Aulas', 'Remarketing Vendas', 'Flash Opening']
        # Mapping rows. Dist=Row 7, Capt=8...
        row_offset = 7
        for i, p in enumerate(phases):
            r = row_offset + i
            invest_map[p] = [
                get_val(df_invest, r, 6), # Total
                get_val(df_invest, r, 8), # Meta
                get_val(df_invest, r, 9), # YT
                get_val(df_invest, r, 10)  # Google
            ]

        # Load Metas por Dia
        df_daily = pd.read_excel(file_path, sheet_name='Metas por Dia', header=None)
        
        header_row = df_daily.iloc[7]
        col_meta_inv = find_col(header_row, 'META', 4)
        col_yt_inv = find_col(header_row, 'YOUTUBE', 4)
        col_g_inv = find_col(header_row, 'GOOGLE', 4)
        col_traf_leads = find_col(header_row, 'TRÁFEGO', 4)
        col_social = find_col(header_row, 'SOCIAL', 4)
        col_mailing = find_col(header_row, 'MAILING', 4)

        daily_data = []
        # Iterate rows 8 to 38 (Captação period)
        for r in range(8, 39):
            date_val = get_val(df_daily, r, 2)
            if pd.isna(date_val): continue
            
            # Meta
            m_inv = get_val(df_daily, r, col_meta_inv)
            m_leads = clean_int(get_val(df_daily, r, col_meta_inv + 1))
            
            # YT
            yt_inv = get_val(df_daily, r, col_yt_inv)
            yt_leads = clean_int(get_val(df_daily, r, col_yt_inv + 1))
            
            # Google
            g_inv = get_val(df_daily, r, col_g_inv)
            
            # Trafego
            t_leads = clean_int(get_val(df_daily, r, col_traf_leads))
            
            # Social
            s_leads = clean_int(get_val(df_daily, r, col_social))
            
            # Mailing
            mail_leads = clean_int(get_val(df_daily, r, col_mailing))
            
            pct_val = get_val(df_daily, r, col_g_inv + 2)
            
            daily_data.append({
                'Fase': 'Captação',
                'Data': date_val,
                '[Investimento] Meta Ads': m_inv,
                '[Leads] Meta Ads': m_leads,
                '[Investimento] YouTube Ads': yt_inv,
                '[Leads] YouTube Ads': yt_leads,
                '[Investimento] Google Ads': g_inv,
                '[Leads] Google Ads': clean_int(get_val(df_daily, r, col_g_inv + 1)), # Leads next to inv
                '%': pct_val * 100 if pct_val is not None else None, # % next to leads
                '[Leads] Tráfego': t_leads,
                '[Leads] Social': s_leads,
                '[Leads] Mailing': mail_leads
            })

        df_daily_out = pd.DataFrame(daily_data)


        with pd.ExcelWriter(os.path.join(OUTPUT_DIR, file3_name)) as writer:
            # Sheet: Meta de Ingressos (Template)
            pd.DataFrame({
                'Fonte & Categoria': ['Tráfego', 'Meta Ads', 'YouTube Ads', 'Google Ads', 'Social', 'Mailing'],
                'Meta': [None]*6
            }).to_excel(writer, sheet_name='Meta de Ingressos', index=False)
            
            # Sheet: Meta de Leads
            pd.DataFrame({
                'Fonte & Categoria': ['Tráfego', 'Meta Ads', 'YouTube Ads', 'Google Ads', 'Social', 'Mailing'],
                'Meta': [
                    leads_data['Tráfego'], leads_data['Meta Ads'], leads_data['YouTube Ads'],
                    leads_data['Google Ads'], leads_data['Social'], leads_data['Mailing']
                ]
            }).to_excel(writer, sheet_name='Meta de Leads', index=False)
            
            # Sheet: Meta de Investimento
            inv_rows = []
            for p in phases:
                vals = invest_map.get(p, [0,0,0,0])
                inv_rows.append([p] + vals)
            
            df_inv = pd.DataFrame(inv_rows, columns=['Fase', '[Previsto] Total', '[Previsto] Meta Ads', '[Previsto] YouTube Ads', '[Previsto] Google Ads'])
            df_inv.to_excel(writer, sheet_name='Meta de Investimento', index=False)
            
            # Sheet: Verba Prevista (Clone of Investimento for now + extra cols if needed)
            df_verba = df_inv.copy()
            df_verba.columns = ['Fase', '[Verba] Total', '[Verba] Meta Ads', '[Verba] YouTube Ads', '[Verba] Google Ads']
            df_verba.to_excel(writer, sheet_name='Verba Prevista', index=False)
            
            # Sheet: Metas por Dia
            df_daily_out.to_excel(writer, sheet_name='Metas por Dia', index=False)
            
            # Sheet: Contas de Anúncio
            pd.DataFrame({
                'Plataforma': ['Meta Ads', 'Google Ads'],
                'Conta de Anúncios': [account_meta, account_google]
            }).to_excel(writer, sheet_name='Contas de Anúncio', index=False)
            
        print(f"  Generated 3 files for {campaign_id}")

    except Exception as e:
        print(f"  ERROR processing {file_path}: {e}")

# Main execution
files = glob.glob(os.path.join(SOURCE_DIR, "*.xlsx"))
print(f"Found {len(files)} files in {SOURCE_DIR}")

for f in files:
    # Skip temporary excel files (start with ~$
    if os.path.basename(f).startswith("~$"):
        continue
    process_file(f)