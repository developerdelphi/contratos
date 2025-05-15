# utils/contract_processing.py
import datetime
from num2words import num2words
import os # Adicionado para os.path.exists e os.makedirs, embora não usado diretamente aqui, mas pode ser útil para funções relacionadas

# Não precisamos de 'app' ou 'current_app' aqui se passarmos o logger

def preparar_dados_para_contrato(selected_donatario_data, valor_bruto_doacao, aliquota_percentual, app_logger, cidade_doador_fixo="Mossoró", uf_doador_fixo="RN"):
    """
    Prepara o dicionário de substituições para o contrato.
    Calcula impostos, valores líquidos e formata dados.
    Retorna um dicionário de substituições.
    """
    app_logger.debug(f"Preparando dados para contrato: Donatário={selected_donatario_data.get('NOME')}, Valor Bruto={valor_bruto_doacao}, Alíquota={aliquota_percentual}%")
    
    # Cálculos
    valor_itcmd = (valor_bruto_doacao * aliquota_percentual) / 100.0
    valor_liquido_doacao = valor_bruto_doacao - valor_itcmd

    # Data por extenso
    data_atual = datetime.date.today()
    meses_extenso = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", 
                     "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    try:
        data_formatada_extenso = f"{cidade_doador_fixo}/{uf_doador_fixo}, {data_atual.day} de {meses_extenso[data_atual.month - 1]} de {data_atual.year}"
    except IndexError:
        app_logger.warning("Mês inválido para formatação de data por extenso, usando formato numérico.")
        data_formatada_extenso = f"{cidade_doador_fixo}/{uf_doador_fixo}, {data_atual.strftime('%d/%m/%Y')}"

    # Valores por extenso
    valor_bruto_extenso = "[ERRO NA GERAÇÃO POR EXTENSO]"
    valor_itcmd_extenso = "[ERRO NA GERAÇÃO POR EXTENSO]"
    valor_liquido_extenso = "[ERRO NA GERAÇÃO POR EXTENSO]"
    try:
        valor_bruto_extenso = num2words(valor_bruto_doacao, lang='pt_BR', to='currency')
        valor_itcmd_extenso = num2words(valor_itcmd, lang='pt_BR', to='currency')
        valor_liquido_extenso = num2words(valor_liquido_doacao, lang='pt_BR', to='currency')
    except Exception as e_num2words:
        app_logger.error(f"Erro ao converter números para extenso com num2words: {e_num2words}", exc_info=True)
        # Mantém o valor de erro definido acima para que o contrato possa ser gerado e o erro notado

    # !! ADAPTE AS CHAVES DE selected_donatario_data PARA CORRESPONDER ÀS SUAS COLUNAS !!
    substituicoes = {
        "NOME_DONATARIO": str(selected_donatario_data.get('NOME', '')).strip().upper(),
        "NACIONALIDADE_DONATARIO": str(selected_donatario_data.get('NACIONALIDADE', 'N/D')).strip().lower(),
        "ESTADO_CIVIL_DONATARIO": str(selected_donatario_data.get('ESTADO_CIVIL', 'N/D')).strip().lower(),
        "PROFISSAO_DONATARIO": str(selected_donatario_data.get('PROFISSAO', 'N/D')).strip().lower(),
        "RG_DONATARIO": str(selected_donatario_data.get('RG', 'N/D')).strip().upper(),
        "CPF_DONATARIO": str(selected_donatario_data.get('CPF', '')).strip(),
        "ENDERECO_DONATARIO": str(selected_donatario_data.get('ENDERECO', 'N/D')).strip().lower().capitalize(),
        "CIDADE_UF_DONATARIO": str(selected_donatario_data.get('CIDADE_UF', 'N/D')).strip().capitalize(),
        "CEP_DONATARIO": str(selected_donatario_data.get('CEP', 'N/D')).strip(),
        "TELEFONE_DONATARIO": str(selected_donatario_data.get('TELEFONE', 'N/D')).strip(),
        "EMAIL_DONATARIO": str(selected_donatario_data.get('EMAIL', 'N/D')).strip().lower(),
        "BANCO_DONATARIO": str(selected_donatario_data.get('BANCO', 'N/D')).strip().upper(),
        "AGENCIA_DONATARIO": str(selected_donatario_data.get('AGENCIA', 'N/D')).strip().upper(),
        "CONTA_DONATARIO": str(selected_donatario_data.get('CONTA', 'N/D')).strip().upper(),
        "CONTA_TIPO": str(selected_donatario_data.get('OPERACAO', 'N/D')).strip().lower(),

        "VALOR_BRUTO_DOACAO_NUM": f"{valor_bruto_doacao:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
        "VALOR_BRUTO_DOACAO_EXTENSO": valor_bruto_extenso.upper(),
        "ALIQUOTA_ITCMD_PERCENTUAL": f"{aliquota_percentual:,.2f}%".replace('.', ','),
        "VALOR_ITCMD_NUM": f"{valor_itcmd:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
        "VALOR_ITCMD_EXTENSO": valor_itcmd_extenso.upper(),
        "VALOR_LIQUIDO_DOACAO_NUM": f"{valor_liquido_doacao:,.2f}".replace('.', '#').replace(',', '.').replace('#', ','),
        "VALOR_LIQUIDO_DOACAO_EXTENSO": valor_liquido_extenso.upper(),
        "LOCAL_DATA_COMPLETA": data_formatada_extenso.upper(),
    }
    app_logger.debug(f"Dicionário de substituições preparado: {substituicoes}")
    return substituicoes

    """
    Carrega um template DOCX, substitui os placeholders e retorna o objeto Document modificado.
    Levanta FileNotFoundError se o template não for encontrado, ou outras exceções.
    """
    app_logger.debug(f"Preenchendo template DOCX: {caminho_template}")
    try:
        if not os.path.exists(caminho_template):
            app_logger.error(f"Arquivo de template DOCX não encontrado em: {caminho_template}")
            raise FileNotFoundError(f"O arquivo de template '{caminho_template}' não foi encontrado.")

        document = Document(caminho_template)
        
        # Substituição em parágrafos
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                text = run.text
                modified = False
                for key, value in dados_substituicao.items():
                    if key in text:
                        text = text.replace(key, str(value))
                        modified = True
                if modified:
                    run.text = text
        
        # Substituição em tabelas
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            text = run.text
                            modified = False
                            for key, value in dados_substituicao.items():
                                if key in text:
                                    text = text.replace(key, str(value))
                                    modified = True
                            if modified:
                                run.text = text
        app_logger.debug("Template DOCX preenchido com sucesso.")
        return document
    except FileNotFoundError as fnf_error: # Re-levanta FileNotFoundError especificamente
        app_logger.error(f"FileNotFoundError ao tentar abrir/preencher o template DOCX '{caminho_template}': {fnf_error}", exc_info=True)
        raise
    except Exception as e:
        app_logger.error(f"Erro genérico ao preencher o template DOCX '{caminho_template}': {e}", exc_info=True)
        raise # Re-levanta outras exceções para serem tratadas pela rota