from docx import Document as Docx
from datetime import datetime
import pythoncom
from src.find_replace import Find_Replace
from src.functions import Doc_Rtf, Receita 
from src.functions import find_and_update_table
from src.functions import remove_paginas_vazias_rapido, copiar_plano_de_acao, copiar_inventario_via_range, exportar_para_pdf, atualizar_indice, remove_rows_with_text, highlight_cells_with_text, replace_text_with_images, formatar_e_inserir_conteudo_direto

from src.settings import TOKEN
import traceback
import time, sys
import os
import re

meses = {'JANEIRO':1, 
         'FEVEREIRO':2, 
         'MARÇO':3, 
         'ABRIL':4, 
         'MAIO':5, 
         'JUNHO':6, 
         'JULHO':7, 
         'AGOSTO':8, 
         'SETEMBRO':9, 
         'OUTUBRO':10, 
         'NOVEMBRO':11, 
         'DEZEMBRO':12}

def main(file_base_rtf:str, pgr_modelo:str, pgr_destino:str, pdf_path:str):
    hoje = datetime.now().strftime("%d-%m-%Y %H:%M")
    pythoncom.CoInitialize()
    try:
        print("\nExtraindo dados do PGR arquivo de entrada...")
        rtf = Doc_Rtf(file_base_rtf)
        find_replace = Find_Replace(pgr_modelo)
        cnpj_numeros = str(rtf.cnpj).replace('.','').replace('/','').replace('-','').strip()

        print("Consultando empresa na receita federal...")
        json_receita = Receita().consulta_cnpj_receita_federal(cnpj_numeros, TOKEN)
        keys_receita = Receita().extrair_dados_receita(json_receita)

        # Concatena as informações do doc RTF, com as informações da receita federal
        keys_values = {**rtf.keys_rtf, **keys_receita}

        datas_arquivo = keys_values['ref_data_vigencia'].split()
        nome_arquivo = datas_arquivo[0] + ' - ' + datas_arquivo[1] + ' - ' + datas_arquivo[4] + ' - PGR - ' + keys_values['ref_nome_empresa'] 
        pgr_destino = pgr_destino.replace('nome_arquivo_novo', nome_arquivo)
        
        print("Localizando e substituindo texto no arquivo destino...")
        
        var_grauRisco = keys_values['ref_grau_risco']
        
        grau_risco_map = {
            "1": ['X', 'X', 'X', 'X', 'X', 'X', '1', '1', '1', '1', '1', '1', '1', '1', '2', '2', '4', '3', '5', '4', '6', '5', '8', '6', '1', '1'],
            "2": ['X', 'X', 'X', 'X', '1', '1', '1', '1', '2', '1', '2', '1', '3', '2', '4', '3', '5', '4', '6', '5', '8', '6', '10', '8', '1', '1'],
            "3": ['1', '1', '1', '1', '2', '1', '2', '1', '2', '1', '3', '2', '4', '2', '5', '4', '6', '4', '8', '6', '10', '8', '12', '8', '2', '2'],
            "4": ['1', '1', '2', '1', '3', '2', '3', '2', '4', '2', '4', '2', '4', '3', '5', '4', '6', '5', '9', '7', '11', '8', '13', '10', '2', '2']
        }

        if var_grauRisco in grau_risco_map:
            varCol31, varCol32, varCol41, varCol42, varCol51, varCol52, varCol61, varCol62, varCol71, varCol72, \
            varCol81, varCol82, varCol91, varCol92, varCol101, varCol102, varCol111, varCol112, varCol121, varCol122, \
            varCol131, varCol132, varCol141, varCol142, varCol151, varCol152 = grau_risco_map[var_grauRisco]

        replacements = {
            "NOME DA EMPRESA": keys_values['ref_nome_empresa'],
            "XX.XX.XXXX": keys_values['ref_campo_data'],
            "MÊS DE VIGENCIA ANO VIGENCIA A MÊS DE VIGENCIA ANO VIGENCIA": keys_values['ref_data_vigencia'],
            "cartao_cnpj": format_cnpj(cnpj_numeros),
            "cartao_dataAbertura": format_date(keys_values['rec_dt_abertura']),
            "cartao_nome_empresa": keys_values['rec_nome_empresa'],
            "cartao_nomeFantasia": keys_values['rec_nome_fantasia'],
            "cartao_porte": keys_values['rec_porte'],
            "cartao_codigoDescricao": keys_values['rec_cod_desc_mun'],
            "cartao_codigoDescSec": keys_values['rec_cod_desc_secund'],
            "cartao_codigo_desc_nat": keys_values['rec_cod_desc_jur'],
            "cartao_logradouro": keys_values['rec_logradouro'],
            "cartao_numero": keys_values['rec_num'],
            "cartao_complemento": keys_values['rec_complemento'],
            "cartao_cep": keys_values['rec_cep'],
            "cartao_bairro": keys_values['rec_bairro'],
            "cartao_municipio": keys_values['rec_municipio'],
            "cartao_uf": keys_values['rec_uf'],
            "cartao_email": keys_values['rec_mail'],
            "cartao_telefone": keys_values['rec_tel'],
            "cartao_situacao": keys_values['rec_situacao_cad'],
            "cartao_dataSitCadastral": format_date(keys_values['rec_dta_situacao']),
            "varCnae": keys_values['ref_cnae'],
            "varGrauRisco": var_grauRisco,
            "varDescCnae": keys_values['ref_descr_cnae'],
            "qtdFunMas": keys_values['ref_masculino'],
            "qtdFunFem": keys_values['ref_feminino'],
            "qtdFunMenor": keys_values['ref_menor'],
            "qtdFunTotal": keys_values['ref_total'],
            "CURITIBA/PR, 00 de Abril de 2024": get_current_date(),
            "varGrauRisc": var_grauRisco,
            "varCol31": varCol31,
            "varCol32": varCol32,
            "varCol41": varCol41,
            "varCol42": varCol42,
            "varCol51": varCol51,
            "varCol52": varCol52,
            "varCol61": varCol61,
            "varCol62": varCol62,
            "varCol71": varCol71,
            "varCol72": varCol72,
            "varCol81": varCol81,
            "varCol82": varCol82,
            "varCol91": varCol91,
            "varCol92": varCol92,
            "varCol101": varCol101,
            "varCol102": varCol102,
            "varCol111": varCol111,
            "varCol112": varCol112,
            "varCol121": varCol121,
            "varCol122": varCol122,
            "varCol131": varCol131,
            "varCol132": varCol132,
            "varCol141": varCol141,
            "varCol142": varCol142,
            "varCol151": varCol151,
            "varCol152": varCol152
        }
        
        for find, repl in replacements.items():
            if "None" in repl: repl = "*****"
            find_replace.replace_text(find, repl) 
            find_replace.replace_in_paragraphs(find, repl)
            find_replace.replace_in_shapes(find, repl)
            find_replace.replace_in_headers_and_footers(find, repl)
            
        print(f"Salvando arquivo de saida fase 1")
        
        #Alterar para o nome correto do arquivo de saida
        
        find_replace.save_close_file(pgr_destino)
        time.sleep(3)

        docx = Docx(pgr_destino)
        print("Removendo linhas de planos de ações desnecessários")

        if rtf.prog_auditivo == False:
            remove_rows_with_text(docx, "AUDITIVA:")
        if rtf.prog_respiratorio == False:
            remove_rows_with_text(docx, "RESPIRATÓRIA:")
        if rtf.prog_NR6 == False:
            remove_rows_with_text(docx, "NR 06")
        if rtf.prog_NR10 == False:
            remove_rows_with_text(docx, "NR 10")
        if rtf.prog_NR11 == False:
            remove_rows_with_text(docx, "NR 11")
        if rtf.prog_NR12 == False:
            remove_rows_with_text(docx, "NR 12")
        if rtf.prog_NR13 == False:
            remove_rows_with_text(docx, "NR 13")
        if rtf.prog_NR33 == False:
            remove_rows_with_text(docx, "NR 33")
        if rtf.prog_NR35 == False:
            remove_rows_with_text(docx, "NR 35")

        # Marcar X na posição do mes, na linha NR 01: ELABORAÇÃO DO GRO
        mes_str = str(rtf.data_vigencia.split()[0])
        mes_int = meses[mes_str]
        agora = hoje[:10].replace('-','.') + " Clinimerces"
        find_and_update_table(docx, "NR 01: ELABORAÇÃO DO GRO", "x", mes_int, agora)
        time.sleep(2)

        print("Pintando as celulas da tabela membros da CIPA")
        
        highlight_cells_with_text(docx, rtf.ref_cipa)

        print(f"Encerrando e salvando arquivo resultado em\n{pgr_destino}")
        
        #Remover o texto do item 18 e adiciona os print do pdf
        replace_text_with_images(pdf_path, docx)
        docx.save(pgr_destino)
            
        pythoncom.CoUninitialize()

        time.sleep(5)
        # formatar arquivo rft e copiar para docx de saida
        # TODO: VALIDAR SE VAI PRECISAR MANTER
        # formatar_e_inserir_conteudo_direto(file_base_rtf, pgr_destino)
        time.sleep(5)
        
        pgr_destino = copiar_plano_de_acao(file_base_rtf, pgr_destino)
       
        time.sleep(5)
    
        pgr_destino = copiar_inventario_via_range(file_base_rtf, pgr_destino, keys_values['ref_nome_empresa'])
       
        time.sleep(5)
       
        print("Atualizando Indices")
        atualizar_indice(pgr_destino)
        
        time.sleep(5)
        
        print("Exportando para PDF")
        
        pdf_destino = pgr_destino.replace(".docx", ".pdf")
        exportar_para_pdf(pgr_destino, pdf_destino)
        
        # gravar log de sucesso
        msg_log = f'''{hoje};{rtf.cnpj};sucesso\n'''
        with open('log_de_sucesso.csv', "a") as f:
            f.write(msg_log)
    except Exception as e:
        print(e)
        tb = traceback.format_exc(limit=2)
        try: 
            msg_log = f'''\n{hoje} ######################\n
            Erro no PRG da empresa: CNPJ: {rtf.cnpj}\n
            Descr. erro: {tb}\n'''
        except:
            msg_log = f'''\n{hoje} ######################\n
            Descr. erro: {tb}\n'''
        with open('log_de_erro.csv', "a") as f:
            f.write(msg_log)

def format_date(date_str):
    try:
        date_obj = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return date_obj.strftime('%d/%m/%Y')
    except ValueError:
        return 'Data inválida'

def format_cnpj(cnpj):
    return re.sub(r"(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})", r"\1.\2.\3/\4-\5", cnpj)
    

def get_current_date():
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    
    hoje = datetime.today()
    return f"CURITIBA, {hoje.day} de {meses[hoje.month]} de {hoje.year}"
 
# Conditional block to run the script directly
if __name__ == "__main__":
    USERNAME = os.getenv("USERNAME")
    path_folder = fr'C:\Users\{USERNAME}\Desktop\pgr\files'
    file_base_rtf = fr"C:\\Users\\{USERNAME}\\Desktop\pgr\\files\\2024 2026 - PGR - 76710649000249 - CLUB ATHLETICO PARANAENSE.rtf"
    # file_base_rtf = "C:\\Temp\\files\\2024 2026 - PGR - 06905926000102 - BRASILPACK INDUSTRIA E COMECIO DE EMBALAGENS LTDA.rtf"
    pgr_modelo = fr"C:\\Users\\{USERNAME}\\Desktop\pgr\\files\\PGR_modelo.docx"
    pgr_destino = fr"C:\\Users\\{USERNAME}\\Desktop\pgr\\files\\PGR_resultado_CLUB ATHLETICO PARANAENSE.docx"

    main(file_base_rtf, pgr_modelo, pgr_destino)
