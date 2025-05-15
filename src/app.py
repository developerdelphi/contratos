import os
import pandas as pd 
import logging
import datetime

# Imports do Flask e extensões
from flask import (Flask, render_template, request, redirect, url_for, 
                   flash, session, send_from_directory)
from flask_sqlalchemy import SQLAlchemy
from flask_session import Session

# Imports para novo fluxo de PDF
from jinja2 import Environment, FileSystemLoader
import markdown2 # Para converter Markdown em HTML
from weasyprint import HTML, CSS # Para converter HTML em PDF

# Imports dos seus módulos utils
from utils.google_services import get_sheet_data 
from utils.contract_processing import preparar_dados_para_contrato
# Outros imports que você tinha
from num2words import num2words # Usado em preparar_dados_para_contrato
from config import Config # Sua classe de configuração

# Configuração inicial do Flask
app = Flask(__name__)
app.config.from_object(Config) # Carrega configurações do config.py
app.template_folder = 'templates' # Onde estão index.html, sucesso_geracao.html
app.static_folder = 'static'    # Onde podem estar CSS, JS, imagens
app.logger.setLevel(logging.DEBUG)

# Configurações que vêm do Config (UPLOAD_FOLDER, CREDENTIALS_FILE)
# Elas serão acessadas via app.config['CHAVE'] ou app.config.get('CHAVE')

# Inicializa o SQLAlchemy
db = SQLAlchemy(app)
app.config['SESSION_SQLALCHEMY'] = db # Diz ao Flask-Session para usar esta instância db

# Inicializa a extensão Flask-Session
server_session = Session(app)

# --- Criação da Tabela de Sessões ---
with app.app_context():
    db.create_all()
    app.logger.info("Tabela de sessões verificada/criada no banco de dados SQLite.")

# Configurar o ambiente Jinja2 para carregar templates Markdown
# Assumindo que app.py está em /app e templates em /app/templates/
markdown_template_dir = os.path.join(os.path.dirname(__file__), 'templates', 'markdown_templates')
jinja_markdown_env = Environment(
    loader=FileSystemLoader(markdown_template_dir),
    autoescape=True # Autoescape é bom para HTML, pode não ser estritamente necessário para MD->HTML
)
app.logger.info(f"Ambiente Jinja2 para Markdown configurado para carregar de: {markdown_template_dir}")


@app.route('/', methods=['GET', 'POST'])
def index():
    app.logger.info("--- FUNÇÃO INDEX ACESSADA (FLUXO MARKDOWN) ---")
    
    donatarios_df = None
    # path_template_docx não é mais necessário aqui
    donatarios_para_exibicao = []

    sheet_url_value = session.get('latest_sheet_url', '')
    # doc_url_value não é mais necessário se o template for local

    if request.method == 'POST':
        sheet_url_from_form = request.form.get('sheet_url')
        # doc_url_from_form não é mais pego do formulário
        
        sheet_url_value = sheet_url_from_form
        session['latest_sheet_url'] = sheet_url_from_form
        # session['latest_doc_url'] não é mais necessário

        if not sheet_url_from_form: # Apenas a URL da planilha é obrigatória agora
            flash("Por favor, forneça a URL/ID da Planilha Google.", "danger")
            session.pop('donatarios_data_json', None)
        else:
            app.logger.info(f"Carregando dados da planilha: {sheet_url_from_form}")
            # A função get_sheet_data em utils.google_services precisa ser ajustada
            # para receber app.config['CREDENTIALS_FILE'] ou o caminho diretamente.
            # Por agora, vamos assumir que ela consegue pegar de Config importado no módulo dela.
            donatarios_df = get_sheet_data(sheet_url_from_form, app.logger)

            if donatarios_df is not None:
                app.logger.info(f"donatarios_df carregado com {len(donatarios_df)} linhas.")
                session['donatarios_data_json'] = donatarios_df.to_json(orient='records')
                app.logger.info("Dados completos dos donatários salvos na sessão.")
                flash("Planilha carregada com sucesso!", "success") # Mensagem simplificada
            else:
                app.logger.warning("Falha ao carregar donatarios_df da planilha.")
                session.pop('donatarios_data_json', None)
                # flash já deve ter sido chamado por get_sheet_data em caso de erro
    
    # Lógica comum para GET e para continuar após POST
    if donatarios_df is None and 'donatarios_data_json' in session:
        app.logger.info("Tentando carregar donatarios_df da sessão.")
        try:
            donatarios_df = pd.read_json(session['donatarios_data_json'], orient='records')
            app.logger.info(f"donatarios_df carregado da sessão com {len(donatarios_df)} linhas.")
        except Exception as e:
            app.logger.error(f"Erro ao ler 'donatarios_data_json' da sessão: {e}", exc_info=True)
            session.pop('donatarios_data_json', None)
            donatarios_df = None

    # Processa donatarios_df para exibição
    if donatarios_df is not None and not donatarios_df.empty:
        app.logger.info("Processando donatarios_df para criar donatarios_para_exibicao.")
        nome_coluna_nome = app.config.get('COLUNA_NOME_PADRAO', 'NOME') # Exemplo de pegar de config
        nome_coluna_cpf = app.config.get('COLUNA_CPF_PADRAO', 'CPF')

        colunas_identificadas = []
        if nome_coluna_nome in donatarios_df.columns: colunas_identificadas.append(nome_coluna_nome)
        else: app.logger.warning(f"ALERTA: Coluna '{nome_coluna_nome}' NÃO encontrada na planilha.")
        if nome_coluna_cpf in donatarios_df.columns: colunas_identificadas.append(nome_coluna_cpf)
        else: app.logger.warning(f"ALERTA: Coluna '{nome_coluna_cpf}' NÃO encontrada na planilha.")
        
        if colunas_identificadas:
            donatarios_para_exibicao = donatarios_df[colunas_identificadas].to_dict(orient='records')
            app.logger.info(f"donatarios_para_exibicao preparado com {len(donatarios_para_exibicao)} registros.")
        else:
            app.logger.error("ERRO: Nenhuma das colunas (NOME/CPF) foi encontrada. Tabela vazia.")
            donatarios_para_exibicao = []
    else:
        donatarios_para_exibicao = []

    app.logger.info(f"Antes de renderizar index: len(donatarios_para_exibicao) = {len(donatarios_para_exibicao)}")
    return render_template('index.html', 
                           donatarios_exibicao=donatarios_para_exibicao,
                           sheet_url_value=sheet_url_value)


@app.route('/gerar_contrato', methods=['POST'])
def gerar_contrato():
    app.logger.info("--- ROTA GERAR_CONTRATO (FLUXO MARKDOWN/WEASYPRINT) ---")

    selected_donatario_index_str = request.form.get('donatario_selecionado_index')
    valor_doacao_str = request.form.get('valor_doacao')
    aliquota_str = request.form.get('aliquota')
    # doc_template_path não é mais necessário do formulário

    app.logger.debug(f"Formulário recebido - Índice Str: {selected_donatario_index_str}, Valor Str: {valor_doacao_str}, Alíquota Str: {aliquota_str}")

    if not all([selected_donatario_index_str, valor_doacao_str, aliquota_str]):
        flash("Dados incompletos para gerar o contrato.", "danger")
        return redirect(url_for('index'))

    try:
        selected_donatario_index = int(selected_donatario_index_str)
        valor_bruto_doacao = float(valor_doacao_str)
        aliquota_percentual = float(aliquota_str)
    except (ValueError, TypeError) as e_conv:
        flash("Valores inválidos ou ausentes para os campos do formulário.", "danger")
        app.logger.error(f"Erro ao converter dados do formulário: {e_conv}", exc_info=True)
        return redirect(url_for('index'))

    donatarios_json_str = session.get('donatarios_data_json')
    if not donatarios_json_str:
        flash("Sessão expirada ou dados dos donatários não encontrados.", "warning")
        return redirect(url_for('index'))

    try:
        all_donatarios_list = pd.read_json(donatarios_json_str, orient='records').to_dict(orient='records')
    except Exception as e_json:
        flash("Erro ao processar dados dos donatários da sessão.", "danger")
        app.logger.error(f"Erro ao ler donatarios_data_json: {e_json}", exc_info=True)
        return redirect(url_for('index'))

    if not (0 <= selected_donatario_index < len(all_donatarios_list)):
        flash("Índice de donatário selecionado inválido.", "danger")
        return redirect(url_for('index'))
    
    selected_donatario_data = all_donatarios_list[selected_donatario_index]
    app.logger.info(f"Dados do Donatário: {selected_donatario_data.get('NOME')}")

    try:
        contexto_contrato = preparar_dados_para_contrato(
            selected_donatario_data, valor_bruto_doacao, aliquota_percentual, app.logger
        )

        template_md = jinja_markdown_env.get_template('CONTRATO_MODELO_DOACAO.md.jinja') # Nome do seu arquivo
        markdown_renderizado = template_md.render(contexto_contrato)
        app.logger.debug("Template Markdown renderizado com Jinja2.")

        html_content = markdown2.markdown(markdown_renderizado, extras=["tables", "smarty-pants", "cuddled-lists", "footnotes"])
        app.logger.debug("Markdown convertido para HTML.")
        # app.logger.debug(f"HTML Gerado (primeiros 500 chars): {html_content[:500]}") # Loga o início do HTML
        # Salvar HTML para depuração
        # with open(os.path.join(app.config.get('UPLOAD_FOLDER', 'uploads'), 'debug_contrato.html'), 'w', encoding='utf-8') as f_html:
            # f_html.write(html_content)
        
        # app.logger.info("HTML de depuração salvo em uploads/debug_contrato.html")

        # Opcional: Carregar CSS
        # css_filepath = os.path.join(app.static_folder, 'css', 'contrato_estilo.css')
        # contrato_css_obj = CSS(filename=css_filepath) if os.path.exists(css_filepath) else None
        contrato_css_obj = None # Comece sem CSS externo para simplificar

        data_atual_obj = datetime.date.today()
        nome_donatario_arq = "".join(c if c.isalnum() else "_" for c in selected_donatario_data.get('NOME', 'donatario'))
        nome_arquivo_pdf = f"CONTRATO_{nome_donatario_arq}_{data_atual_obj.strftime('%Y%m%d')}.pdf"
        
        pasta_contratos = os.path.join(app.config.get('UPLOAD_FOLDER', 'uploads'), 'contratos_gerados')
        if not os.path.exists(pasta_contratos): os.makedirs(pasta_contratos)
        
        caminho_pdf_final = os.path.join(pasta_contratos, nome_arquivo_pdf)

        app.logger.info(f"Gerando PDF com WeasyPrint para: {caminho_pdf_final}")
        html_doc = HTML(string=html_content, base_url=request.url_root)
        
        css_files = []
        css_filepath = os.path.join(app.static_folder, 'css', 'contrato_estilo.css')
        if os.path.exists(css_filepath):
            css_files.append(CSS(filename=css_filepath))
            app.logger.info(f"CSS '{css_filepath}' carregado para o PDF.")
        else:
            app.logger.warning(f"Arquivo CSS '{css_filepath}' não encontrado. Usando estilos padrão do browser/WeasyPrint.")
        
        if not contrato_css_obj:
            # html_doc.write_pdf(caminho_pdf_final, stylesheets=[contrato_css_obj])
            html_doc.write_pdf(caminho_pdf_final, stylesheets=css_files if css_files else None)
        else:
            html_doc.write_pdf(caminho_pdf_final)
        
        app.logger.info(f"Arquivo PDF gerado: {caminho_pdf_final}")

        flash(f"Contrato PDF para {selected_donatario_data.get('NOME')} gerado com sucesso!", "success")
        return render_template('sucesso_geracao.html',
                               nome_arquivo_pdf=nome_arquivo_pdf,
                               nome_donatario=selected_donatario_data.get('NOME'))

    except Exception as e_geral:
        app.logger.error(f"Erro ao gerar contrato (Markdown/WeasyPrint): {e_geral}", exc_info=True)
        flash(f"Erro ao gerar contrato: {str(e_geral)[:100]}", "danger")
        return redirect(url_for('index'))


@app.route('/download_contrato/<path:filename>')
def download_contrato(filename):
    app.logger.info(f"Download requisitado para: {filename}")
    diretorio_contratos = os.path.join(app.config.get('UPLOAD_FOLDER', 'uploads'), 'contratos_gerados')
    try:
        return send_from_directory(directory=diretorio_contratos, path=filename, as_attachment=True)
    except FileNotFoundError:
        app.logger.error(f"Download falhou: arquivo não encontrado em {diretorio_contratos}/{filename}")
        flash("Arquivo não encontrado para download.", "danger")
        return redirect(url_for('index'))
    except Exception as e:
        app.logger.error(f"Erro no download: {e}", exc_info=True)
        flash(f"Erro interno ao processar download: {e}", "danger")
        return redirect(url_for('index'))

if __name__ == '__main__':
    upload_dir = app.config.get('UPLOAD_FOLDER')
    if upload_dir:
        if not os.path.exists(upload_dir): os.makedirs(upload_dir)
        pasta_contratos_gerados_inicial = os.path.join(upload_dir, 'contratos_gerados')
        if not os.path.exists(pasta_contratos_gerados_inicial): os.makedirs(pasta_contratos_gerados_inicial)
    else:
        app.logger.critical("UPLOAD_FOLDER não configurado. O aplicativo não pode iniciar.")
    
    app.run(host='0.0.0.0', port=8000, debug=True)