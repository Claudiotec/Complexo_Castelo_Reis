import os
import sqlite3
import pandas as pd
import difflib
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, send_file, session, jsonify
from werkzeug.utils import secure_filename
from datetime import datetime
from plyer import notification
import zipfile

# ===== CONFIGURAÇÕES =====
PASTA_ARQUIVOS = "arquivos"
PASTA_DADOS = "dados"
PASTA_RECEBIDOS = "dados_recebidos"
BANCO_DADOS = "registros.db"
SENHA_UPLOAD = "Claudtec2011"
SENHA_ADMIN = "Claudtec2011"  # Senha para acessar os dados recebidos

PORT = int(os.environ.get("PORT", 5000))

for pasta in [PASTA_ARQUIVOS, PASTA_DADOS, PASTA_RECEBIDOS]:
    os.makedirs(pasta, exist_ok=True)

# ===== CRIAR BANCO DE DADOS =====
def init_db():
    conn = sqlite3.connect(BANCO_DADOS)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS conversas_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    datahora TEXT, pergunta TEXT, resposta TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS envios_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    datahora TEXT, nome TEXT, contacto TEXT,
                    comentario TEXT, arquivo TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS feedback_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    datahora TEXT, encarregado TEXT, aluno TEXT,
                    contacto TEXT, mensagem TEXT)""")
    conn.commit()
    conn.close()

init_db()

# ===== FUNÇÃO DE NOTIFICAÇÃO =====
def notificar(titulo, mensagem):
    try:
        notification.notify(title=titulo, message=mensagem, timeout=5)
    except:
        pass

# ===== FUNÇÃO PARA ENCONTRAR O LOGO =====
def encontrar_logo():
    # Procurar por arquivos que podem ser o logo (insignia, logo, emblem, etc.)
    possiveis_nomes = ['insignia', 'logo', 'emblem', 'brasao', 'symbol']
    extensoes = ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg']
    
    for nome in possiveis_nomes:
        for ext in extensoes:
            caminho_arquivo = os.path.join(PASTA_ARQUIVOS, nome + ext)
            if os.path.exists(caminho_arquivo):
                return nome + ext
            
            # Também verificar variações com letras maiúsculas
            caminho_arquivo = os.path.join(PASTA_ARQUIVOS, nome.capitalize() + ext)
            if os.path.exists(caminho_arquivo):
                return nome.capitalize() + ext
    
    # Se não encontrar, retorna None
    return None

# ===== CHATBOT =====
def responder_pergunta(pergunta):
    pergunta_original = pergunta.strip()
    pergunta_lower = pergunta_original.lower()
    todas_perguntas, perguntas_respostas = [], []
    resposta_final = ""

    for nome in os.listdir(PASTA_DADOS):
        if nome.endswith(".xlsx"):
            caminho = os.path.join(PASTA_DADOS, nome)
            try:
                df = pd.read_excel(caminho)
                if df.shape[1] < 2:
                    continue
                perguntas = df.iloc[:, 0].astype(str).str.lower().tolist()
                respostas = df.iloc[:, 1].astype(str).tolist()
                todas_perguntas.extend(perguntas)
                perguntas_respostas.extend(zip(perguntas, respostas))
            except Exception as e:
                print(f"Erro ao ler arquivo {nome}: {e}")
                continue

    match = difflib.get_close_matches(pergunta_lower, todas_perguntas, n=1, cutoff=0.6)
    if match:
        for p, r in perguntas_respostas:
            if p == match[0]:
                resposta_final = r
                break
    else:
        sugestoes = difflib.get_close_matches(pergunta_lower, todas_perguntas, n=3, cutoff=0.4)
        if sugestoes:
            resposta_final = "❓ Não encontrei uma resposta exata.\nTalvez você quis dizer:\n• " + "\n• ".join(sugestoes)
        else:
            resposta_final = "❌ Desculpa, não encontrei nenhuma resposta relacionada."

    # SALVAR EM SQLITE
    conn = sqlite3.connect(BANCO_DADOS)
    c = conn.cursor()
    c.execute("INSERT INTO conversas_log (datahora, pergunta, resposta) VALUES (?, ?, ?)",
              (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), pergunta_original, resposta_final))
    conn.commit()
    conn.close()

    return resposta_final

# ===== FLASK APP =====
app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui_complexo_escolar_2024'  # Necessário para usar sessões

@app.route('/', methods=['GET', 'POST'])
def index():
    arquivos = sorted([f for f in os.listdir(PASTA_ARQUIVOS) if os.path.isfile(os.path.join(PASTA_ARQUIVOS, f))])
    resposta = None
    if request.method == 'POST' and 'mensagem' in request.form:
        resposta = responder_pergunta(request.form.get('mensagem'))
    
    # Buscar o logo
    logo = encontrar_logo()
    
    return render_template('index.html', arquivos=arquivos, resposta=resposta, logo=logo)

# --- ROTA PARA VISUALIZAR ARQUIVOS (PREVIEW) ---
@app.route('/visualizar/<nome_arquivo>')
def visualizar_arquivo(nome_arquivo):
    return send_from_directory(PASTA_ARQUIVOS, nome_arquivo)

# --- UPLOAD PROTEGIDO ---
@app.route('/upload_arquivos', methods=['GET', 'POST'])
def upload_arquivos():
    if request.method == 'POST':
        senha = request.form.get("senha")
        if senha != SENHA_UPLOAD:
            return "Senha incorreta!"
        arquivo = request.files['arquivo']
        if arquivo:
            nome_arquivo = secure_filename(arquivo.filename)
            caminho = os.path.join(PASTA_ARQUIVOS, nome_arquivo)
            arquivo.save(caminho)
            return f"Arquivo {nome_arquivo} enviado com sucesso!"
    return '''
    <form method="post" enctype="multipart/form-data">
        <input type="password" name="senha" placeholder="Senha"><br><br>
        <input type="file" name="arquivo">
        <input type="submit" value="Enviar">
    </form>
    '''

@app.route('/upload', methods=['POST'])
def upload():
    nome = request.form.get("nome")
    contacto = request.form.get("contacto")
    comentario = request.form.get("comentario")
    arquivo = request.files['arquivo']

    if arquivo:
        nome_arquivo = secure_filename(arquivo.filename)
        caminho = os.path.join(PASTA_RECEBIDOS, nome_arquivo)
        arquivo.save(caminho)

        # SALVAR EM SQLITE
        conn = sqlite3.connect(BANCO_DADOS)
        c = conn.cursor()
        c.execute("INSERT INTO envios_log (datahora, nome, contacto, comentario, arquivo) VALUES (?, ?, ?, ?, ?)",
                  (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), nome, contacto, comentario, nome_arquivo))
        conn.commit()
        conn.close()

    return redirect(url_for('index'))

@app.route('/download/<nome_arquivo>')
def download(nome_arquivo):
    return send_from_directory(PASTA_ARQUIVOS, nome_arquivo, as_attachment=True)

@app.route('/feedback', methods=['POST'])
def feedback():
    encarregado = request.form.get("encarregado")
    aluno = request.form.get("aluno")
    contacto = request.form.get("contacto")
    mensagem = request.form.get("mensagem")

    conn = sqlite3.connect(BANCO_DADOS)
    c = conn.cursor()
    c.execute("INSERT INTO feedback_log (datahora, encarregado, aluno, contacto, mensagem) VALUES (?, ?, ?, ?, ?)",
              (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encarregado, aluno, contacto, mensagem))
    conn.commit()
    conn.close()

    return redirect(url_for('index'))

# --- EXPORTAR LOGS PARA EXCEL ---
@app.route('/exportar/<tipo>')
def exportar(tipo):
    conn = sqlite3.connect(BANCO_DADOS)
    if tipo == "conversas":
        df = pd.read_sql_query("SELECT * FROM conversas_log", conn)
        arquivo = "conversas_log.xlsx"
    elif tipo == "envios":
        df = pd.read_sql_query("SELECT * FROM envios_log", conn)
        arquivo = "envios_log.xlsx"
    elif tipo == "feedback":
        df = pd.read_sql_query("SELECT * FROM feedback_log", conn)
        arquivo = "feedback_log.xlsx"
    else:
        return "Tipo inválido!"
    conn.close()
    df.to_excel(arquivo, index=False)
    return send_file(arquivo, as_attachment=True)

# --- BAIXAR TODOS ARQUIVOS RECEBIDOS EM ZIP ---
@app.route('/baixar_recebidos')
def baixar_recebidos():
    zip_path = "dados_recebidos.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, dirs, files in os.walk(PASTA_RECEBIDOS):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)
    return send_file(zip_path, as_attachment=True)

# --- LOGIN PARA VISUALIZAR DADOS RECEBIDOS ---
@app.route('/login_dados', methods=['GET', 'POST'])
def login_dados():
    if request.method == 'POST':
        senha = request.form.get('senha')
        if senha == SENHA_ADMIN:
            session['logado'] = True
            return redirect(url_for('visualizar_dados'))
        else:
            return render_template('login.html', erro="Senha incorreta!")
    return render_template('login.html')

# --- VISUALIZAR DADOS RECEBIDOS (PROTEGIDO) ---
@app.route('/visualizar_dados')
def visualizar_dados():
    if not session.get('logado'):
        return redirect(url_for('login_dados'))
    
    # Listar arquivos da pasta dados_recebidos com informações
    arquivos_info = []
    for f in os.listdir(PASTA_RECEBIDOS):
        caminho = os.path.join(PASTA_RECEBIDOS, f)
        if os.path.isfile(caminho):
            tamanho = os.path.getsize(caminho)
            data_modificacao = datetime.fromtimestamp(os.path.getmtime(caminho))
            arquivos_info.append({
                'nome': f,
                'tamanho': tamanho,
                'data_modificacao': data_modificacao
            })
    
    # Ordenar por data de modificação (mais recente primeiro)
    arquivos_info.sort(key=lambda x: x['data_modificacao'], reverse=True)
    
    # Buscar informações do banco de dados
    conn = sqlite3.connect(BANCO_DADOS)
    df_envios = pd.read_sql_query("SELECT * FROM envios_log ORDER BY id DESC", conn)
    conn.close()
    
    return render_template('dados_recebidos.html', 
                         arquivos=arquivos_info, 
                         envios=df_envios.to_dict('records'))

# --- DOWNLOAD DE ARQUIVO INDIVIDUAL DA PASTA RECEBIDOS ---
@app.route('/download_recebido/<nome_arquivo>')
def download_recebido(nome_arquivo):
    if not session.get('logado'):
        return redirect(url_for('login_dados'))
    return send_from_directory(PASTA_RECEBIDOS, nome_arquivo, as_attachment=True)

# --- SERVIR ARQUIVOS DA PASTA RECEBIDOS PARA VISUALIZAÇÃO ---
@app.route('/dados_recebidos/<nome_arquivo>')
def visualizar_recebido(nome_arquivo):
    if not session.get('logado'):
        return redirect(url_for('login_dados'))
    return send_from_directory(PASTA_RECEBIDOS, nome_arquivo)

# --- LOGOUT ---
@app.route('/logout_dados')
def logout_dados():
    session.pop('logado', None)
    return redirect(url_for('index'))

# --- REMOVER ARQUIVO DA PASTA ARQUIVOS (página principal) ---
@app.route('/remover_arquivo', methods=['POST'])
def remover_arquivo():
    data = request.get_json()
    nome_arquivo = data.get('arquivo')
    
    if nome_arquivo:
        caminho_arquivo = os.path.join(PASTA_ARQUIVOS, nome_arquivo)
        if os.path.exists(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                return jsonify({'success': True})
            except Exception as e:
                print(f"Erro ao remover arquivo: {e}")
                return jsonify({'success': False, 'error': str(e)})
    
    return jsonify({'success': False, 'error': 'Arquivo não encontrado'})

# --- REMOVER ITEM (arquivo ou registro) DA PÁGINA DE DADOS RECEBIDOS ---
@app.route('/remover_item', methods=['POST'])
def remover_item():
    if not session.get('logado'):
        return jsonify({'success': False, 'error': 'Não autorizado'})
    
    data = request.get_json()
    item_id = data.get('id')
    tipo = data.get('tipo')
    
    try:
        if tipo == 'envio':
            # Remover registro do banco de dados e arquivo físico se existir
            conn = sqlite3.connect(BANCO_DADOS)
            c = conn.cursor()
            
            # Primeiro buscar o nome do arquivo associado ao registro
            c.execute("SELECT arquivo FROM envios_log WHERE id = ?", (item_id,))
            resultado = c.fetchone()
            
            if resultado:
                nome_arquivo = resultado[0]
                # Remover arquivo físico se existir
                if nome_arquivo:
                    caminho_arquivo = os.path.join(PASTA_RECEBIDOS, nome_arquivo)
                    if os.path.exists(caminho_arquivo):
                        os.remove(caminho_arquivo)
                
                # Remover registro do banco
                c.execute("DELETE FROM envios_log WHERE id = ?", (item_id,))
                conn.commit()
            
            conn.close()
            return jsonify({'success': True})
            
        elif tipo == 'arquivo':
            # Remover apenas o arquivo físico da pasta dados_recebidos
            caminho_arquivo = os.path.join(PASTA_RECEBIDOS, item_id)
            if os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
                return jsonify({'success': True})
            else:
                return jsonify({'success': False, 'error': 'Arquivo não encontrado'})
        
        else:
            return jsonify({'success': False, 'error': 'Tipo inválido'})
            
    except Exception as e:
        print(f"Erro ao remover item: {e}")
        return jsonify({'success': False, 'error': str(e)})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=False)
