import os
import gspread
import pandas as pd
import io
import logging
import datetime
import subprocess
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from flask_session import Session
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from docx import Document
from num2words import num2words
from config import Config
from utils.google_services import (
    get_google_sheets_client,
    get_google_drive_service,
    get_sheet_data,
    download_google_doc_as_docx
)
from utils.contract_processing import preparar_dados_para_contrato, preencher_contrato_docx


# Configuração inicial do Flask
app = Flask(__name__)
app.config.from_object(Config)
app.template_folder = 'templates'
app.static_folder = 'static'
app.logger.setLevel(logging.DEBUG) # Ou logging.INFO

UPLOAD_FOLDER = app.config['UPLOAD_FOLDER']
CREDENTIALS_FILE = app.config['CREDENTIALS_FILE']

# Inicializa o SQLAlchemy
db = SQLAlchemy(app)
# Diz ao Flask-Session para usar sua instância 'db' do SQLAlchemy
app.config['SESSION_SQLALCHEMY'] = db

# Inicializa a extensão Flask-Session DEPOIS de configurar SESSION_TYPE e outras configs
server_session = Session(app) # Isso substitui a gestão de sessão padrão do Flask

# --- Criação da Tabela de Sessões ---
with app.app_context():
    db.create_all()
    app.logger.info("Tabela de sessões verificada/criada no banco de dados SQLite.")

# --- Fim da Configuração para Flask-Session e SQLAlchemy ---


@app.route('/', methods=['GET', 'POST'])
def index():
    app.logger.info("--- FUNÇÃO INDEX ACESSADA ---")
    donatarios_df = None
    path_template_docx = None
    donatarios_para_exibicao = [] # Inicializa como lista vazia

    sheet_url_value = session.get('latest_sheet_url', '')
    doc_url_value = session.get('latest_doc_url', '')

    if request.method == 'POST':
        sheet_url_from_form = request.form.get('sheet_url')
        doc_url_from_form = request.form.get('doc_url')
        sheet_url_value = sheet_url_from_form
        doc_url_value = doc_url_from_form
        session['latest_sheet_url'] = sheet_url_from_form
        session['latest_doc_url'] = doc_url_from_form

        if not sheet_url_from_form or not doc_url_from_form:
            flash("Por favor, forneça as URLs/IDs da Planilha Google e do Documento Google.", "danger")
            session.pop('donatarios_data_json', None)
            session.pop('latest_doc_template_path', None)
        else:
            app.logger.info(f"Carregando dados da planilha: {sheet_url_from_form}")
            donatarios_df = get_sheet_data(sheet_url_from_form, app.logger)

            if donatarios_df is not None:
                app.logger.info(f"donatarios_df carregado com {len(donatarios_df)} linhas.")
                # ATENÇÃO: O AVISO DE COOKIE GRANDE AINDA OCORRERÁ AQUI!
                session['donatarios_data_json'] = donatarios_df.to_json(orient='records')
                app.logger.info("Dados completos dos donatários salvos na sessão (pode exceder o limite do cookie).")
            else:
                app.logger.warning("Falha ao carregar donatarios_df da planilha.")
                session.pop('donatarios_data_json', None)

            app.logger.info(f"Baixando template do documento: {doc_url_from_form}")
            path_template_docx = download_google_doc_as_docx(doc_url_from_form, app.config['UPLOAD_FOLDER'], app.logger)

            if path_template_docx is not None:
                app.logger.info(f"path_template_docx baixado: {path_template_docx}")
                session['latest_doc_template_path'] = path_template_docx
            else:
                app.logger.warning("Falha ao baixar path_template_docx.")
                session.pop('latest_doc_template_path', None)

            if donatarios_df is not None and path_template_docx is not None:
                flash("Planilha e Documento carregados com sucesso!", "success")
            elif donatarios_df is not None and path_template_docx is None:
                flash("Planilha carregada, mas houve erro ao carregar o Documento.", "warning")
            elif donatarios_df is None and path_template_docx is not None:
                flash("Documento carregado, mas houve erro ao carregar a Planilha.", "warning")
    
    # --- Lógica Comum para GET e para continuar após POST ---
    # Tenta carregar donatarios_df da sessão se não foi carregado no POST atual
    if donatarios_df is None and 'donatarios_data_json' in session:
        app.logger.info("Tentando carregar donatarios_df da sessão (GET ou POST falhou em carregar dados novos).")
        try:
            # ATENÇÃO: SE A SESSÃO FOI TRUNCADA/CORROMPIDA, ISTO PODE FALHAR OU RETORNAR DADOS INCOMPLETOS
            donatarios_df = pd.read_json(session['donatarios_data_json'], orient='records')
            app.logger.info(f"donatarios_df carregado da sessão com {len(donatarios_df)} linhas.")
        except Exception as e:
            app.logger.error(f"Erro ao ler 'donatarios_data_json' da sessão (possivelmente corrompido): {e}")
            session.pop('donatarios_data_json', None)
            donatarios_df = None

    if path_template_docx is None and 'latest_doc_template_path' in session:
        candidate_path = session['latest_doc_template_path']
        if os.path.exists(candidate_path):
            path_template_docx = candidate_path
            app.logger.info(f"path_template_docx carregado da sessão: {path_template_docx}")
        else:
            session.pop('latest_doc_template_path', None)

    # Processa donatarios_df para exibição (SEMPRE que donatarios_df tiver dados)
    if donatarios_df is not None and not donatarios_df.empty:
        app.logger.info("Processando donatarios_df para criar donatarios_para_exibicao.")
        
        nome_coluna_para_nome_do_donatario = 'NOME' # Confirmado pelo usuário
        nome_coluna_para_cpf_do_donatario = 'CPF'   # Confirmado pelo usuário

        colunas_identificadas_para_exibicao = []
        if nome_coluna_para_nome_do_donatario in donatarios_df.columns:
            colunas_identificadas_para_exibicao.append(nome_coluna_para_nome_do_donatario)
        else:
            app.logger.warning(f"ALERTA: Coluna '{nome_coluna_para_nome_do_donatario}' NÃO encontrada na planilha.")
        
        if nome_coluna_para_cpf_do_donatario in donatarios_df.columns:
            colunas_identificadas_para_exibicao.append(nome_coluna_para_cpf_do_donatario)
        else:
            app.logger.warning(f"ALERTA: Coluna '{nome_coluna_para_cpf_do_donatario}' NÃO encontrada na planilha.")
        
        if colunas_identificadas_para_exibicao:
            donatarios_para_exibicao = donatarios_df[colunas_identificadas_para_exibicao].to_dict(orient='records')
            app.logger.info(f"donatarios_para_exibicao preparado com {len(donatarios_para_exibicao)} registros e colunas: {list(donatarios_para_exibicao[0].keys()) if donatarios_para_exibicao else 'VAZIO'}.")
        else:
            app.logger.error("ERRO: Nenhuma das colunas (NOME/CPF) foi encontrada. Tabela de donatarios ficará vazia.")
            donatarios_para_exibicao = [] # Garante que é uma lista vazia
    else:
        app.logger.info("donatarios_df está vazio ou None ao final do processamento. 'donatarios_para_exibicao' será uma lista vazia.")
        donatarios_para_exibicao = [] # Garante que é uma lista vazia se não houver dados

    # LOG FINAL ANTES DE RENDERIZAR:
    app.logger.info(f"Antes de renderizar: len(donatarios_para_exibicao) = {len(donatarios_para_exibicao)}")
    if donatarios_para_exibicao:
         app.logger.info(f"Primeiro item em donatarios_para_exibicao: {donatarios_para_exibicao[0]}")


    return render_template('index.html', 
                           donatarios_exibicao=donatarios_para_exibicao,
                           path_template_docx=path_template_docx,
                           sheet_url_value=sheet_url_value,
                           doc_url_value=doc_url_value)


@app.route('/gerar_contrato', methods=['POST'])
def gerar_contrato():
    app.logger.info("--- ROTA GERAR_CONTRATO ACESSADA ---")

    # 1. Obter e validar dados do formulário
    selected_donatario_index_str = request.form.get('donatario_selecionado_index')
    valor_doacao_str = request.form.get('valor_doacao')
    aliquota_str = request.form.get('aliquota')
    doc_template_path = request.form.get('doc_template_path')

    app.logger.debug(f"Formulário recebido - Índice Str: {selected_donatario_index_str}, Valor Str: {valor_doacao_str}, Alíquota Str: {aliquota_str}, Template: {doc_template_path}")

    if not all([selected_donatario_index_str, valor_doacao_str, aliquota_str, doc_template_path]):
        flash("Dados incompletos recebidos para gerar o contrato. Tente novamente.", "danger")
        app.logger.warning("Dados incompletos no formulário para gerar_contrato.")
        return redirect(url_for('index'))

    try:
        selected_donatario_index = int(selected_donatario_index_str)
        valor_bruto_doacao = float(valor_doacao_str)
        aliquota_percentual = float(aliquota_str)
        app.logger.debug(f"Valores convertidos - Índice: {selected_donatario_index}, Valor Bruto: {valor_bruto_doacao}, Alíquota: {aliquota_percentual}%")
    except (ValueError, TypeError) as e_conv: # Captura ambos os erros de conversão
        flash("Valores inválidos ou ausentes para índice, doação ou alíquota. Use números válidos.", "danger")
        app.logger.error(f"Erro ao converter dados do formulário: {e_conv}", exc_info=True)
        return redirect(url_for('index'))

    # 2. Recuperar dados completos do donatário da sessão
    donatarios_json_str = session.get('donatarios_data_json')
    if not donatarios_json_str:
        flash("Sessão expirada ou dados dos donatários não encontrados. Recarregue a planilha.", "warning")
        app.logger.warning("donatarios_data_json não encontrado na sessão.")
        return redirect(url_for('index'))

    try:
        all_donatarios_list = pd.read_json(donatarios_json_str, orient='records').to_dict(orient='records')
    except Exception as e_json:
        flash("Erro ao processar dados dos donatários da sessão. Recarregue a planilha.", "danger")
        app.logger.error(f"Erro ao fazer pd.read_json de donatarios_data_json: {e_json}", exc_info=True)
        return redirect(url_for('index'))

    if not (0 <= selected_donatario_index < len(all_donatarios_list)):
        flash(f"Índice de donatário selecionado ({selected_donatario_index}) inválido.", "danger")
        app.logger.error(f"Índice de donatário selecionado inválido: {selected_donatario_index}. Lista com {len(all_donatarios_list)} itens.")
        return redirect(url_for('index'))

    selected_donatario_data = all_donatarios_list[selected_donatario_index]
    app.logger.info(f"Dados Completos do Donatário Selecionado: {selected_donatario_data.get('NOME')}")

    # 3. Preparar dados e preencher o template DOCX usando as funções auxiliares
    try:
        # Passando app.logger para as funções auxiliares, ou elas podem usar current_app.logger
        dict_substituicoes = preparar_dados_para_contrato(
            selected_donatario_data, 
            valor_bruto_doacao, 
            aliquota_percentual, 
            app.logger)
        
        documento_modificado = preencher_contrato_docx(
            doc_template_path, 
            dict_substituicoes, 
            app.logger)
        
        if documento_modificado is None: # Se preencher_contrato_docx retornou None devido a um erro interno
            flash("Ocorreu um erro ao gerar o conteúdo do contrato.", "danger")
            return redirect(url_for('index'))

        # 4. Salvar o documento DOCX modificado
        data_atual_obj = datetime.date.today() # Objeto data para nome do arquivo
        nome_donatario_arquivo = "".join(c if c.isalnum() else "_" for c in selected_donatario_data.get('NOME', 'donatario_s_nome'))
        nome_arquivo_docx_final = f"CONTRATO_{nome_donatario_arquivo}_{data_atual_obj.strftime('%Y%m%d')}.docx"
        
        pasta_contratos_gerados = os.path.join(app.config.get('UPLOAD_FOLDER', 'uploads'), 'contratos_gerados')
        if not os.path.exists(pasta_contratos_gerados):
            os.makedirs(pasta_contratos_gerados)
            app.logger.info(f"Pasta '{pasta_contratos_gerados}' criada.")

        caminho_docx_final = os.path.join(pasta_contratos_gerados, nome_arquivo_docx_final)
        documento_modificado.save(caminho_docx_final)
        app.logger.info(f"Contrato DOCX modificado salvo em: {caminho_docx_final}")

        # 5. Converter DOCX para PDF
        nome_arquivo_pdf = nome_arquivo_docx_final.replace(".docx", ".pdf")
        caminho_pdf_final = os.path.join(pasta_contratos_gerados, nome_arquivo_pdf)
        output_dir_for_pdf = pasta_contratos_gerados

        app.logger.info(f"Iniciando conversão de '{caminho_docx_final}' para PDF em '{output_dir_for_pdf}'...")
        comando = [ # Use o comando que está funcionando para você (libreoffice ou unoconv)
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", output_dir_for_pdf, caminho_docx_final
        ]
        app.logger.debug(f"Executando comando de conversão PDF: {' '.join(comando)}")
        process = subprocess.run(comando, timeout=60, capture_output=True, text=True, check=False)

        app.logger.debug(f"Conversão PDF finalizada. Código de retorno: {process.returncode}")
        app.logger.debug(f"Conversão PDF stdout: {process.stdout}")
        app.logger.debug(f"Conversão PDF stderr: {process.stderr}")

        if process.returncode == 0 and os.path.exists(caminho_pdf_final):
            app.logger.info(f"Arquivo PDF gerado com sucesso e encontrado: {caminho_pdf_final}")
            flash(f"Contrato PDF para {selected_donatario_data.get('NOME')} gerado com sucesso!", "success")
            return render_template('sucesso_geracao.html', 
                                   nome_arquivo_pdf=nome_arquivo_pdf,
                                   nome_donatario=selected_donatario_data.get('NOME'))
        else:
            app.logger.error(f"Falha na condição de sucesso pós-conversão PDF. Código de retorno: {process.returncode}. PDF existe? {os.path.exists(caminho_pdf_final)}")
            mensagem_erro_flash = "Contrato DOCX gerado, mas problema na conversão para PDF."
            if process.stderr: mensagem_erro_flash += f" Detalhe: {process.stderr[:100]}"
            flash(mensagem_erro_flash, "danger")
            return redirect(url_for('index'))

    except FileNotFoundError as e_fnf: # Erro se o template DOCX não for encontrado
        app.logger.error(f"Erro de arquivo não encontrado ao processar contrato: {e_fnf}", exc_info=True)
        flash(f"Erro crítico: Arquivo template '{doc_template_path}' não encontrado ou inacessível.", "danger")
        return redirect(url_for('index'))
    except Exception as e_geral:
        app.logger.error(f"Erro inesperado na rota gerar_contrato: {e_geral}", exc_info=True)
        flash(f"Ocorreu um erro inesperado ao gerar o contrato: {str(e_geral)[:100]}", "danger")
        return redirect(url_for('index'))
    

@app.route('/download_contrato/<path:filename>')
def download_contrato(filename):
    app.logger.info(f"Requisição de download para o arquivo: {filename}")
    # O diretório base para os contratos gerados (dentro da pasta de uploads)
    diretorio_contratos = os.path.join(app.config['UPLOAD_FOLDER'], 'contratos_gerados')
    app.logger.info(f"Tentando servir arquivo de: {diretorio_contratos}, arquivo: {filename}")
    try:
        return send_from_directory(directory=diretorio_contratos,
                                   path=filename, # Use 'path' em vez de 'filename' como argumento aqui
                                   as_attachment=True) # as_attachment=True força o download
    except FileNotFoundError:
        app.logger.error(f"Arquivo de contrato para download NÃO ENCONTRADO: {os.path.join(diretorio_contratos, filename)}")
        flash("Erro: Arquivo de contrato não encontrado para download.", "danger")
        return redirect(url_for('index'))
    except Exception as e:
        app.logger.error(f"Erro ao tentar enviar arquivo para download: {e}", exc_info=True)
        flash(f"Erro interno ao processar o download: {e}", "danger")
        return redirect(url_for('index'))
    
    
def start_server(host='0.0.0.0', port=8080): # Adicionei os parâmetros aqui
    # Removi o return daqui, app.run() bloqueia.
    app.run(host=host, port=port, debug=True)

if __name__ == '__main__':
    # Certifique-se que a pasta de uploads existe no início
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    start_server(port=8000) # Use a porta que preferir


def start_server(host='0.0.0.0', port=8080):
    return app.run(host=host, 
                   port=port, debug=True)