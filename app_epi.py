from flask import Flask, render_template, request, redirect, url_for, send_file, make_response, session, abort, flash, jsonify
import sqlite3
from datetime import datetime
import os
from io import BytesIO
from openpyxl import Workbook
from barcode import Code39
from barcode.writer import ImageWriter
import barcode
from functools import wraps
import secrets

app = Flask(__name__)

# sessão e credenciais fixas
app.secret_key = "troque-esta-string-por-uma-aleatoria"  # ex.: secrets.token_hex(16)
ADMIN_USER = "admin"
ADMIN_PASS = "Geo@#2025"

def login_required(view):
    """Exige login para qualquer método (GET/POST). Útil p/ editar/excluir."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user"):
            session["next_url"] = request.full_path if request.query_string else request.path
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped

def login_required_for(methods=("POST",)):
    """
    Exige login apenas se a requisição for de um dos métodos informados.
    Ex.: @login_required_for(("POST",)) deixa GET público e protege POST.
    """
    methods = set(m.upper() for m in methods)
    def deco(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if request.method.upper() in methods and not session.get("user"):
                session["next_url"] = request.full_path if request.query_string else request.path
                return redirect(url_for("login"))
            return view(*args, **kwargs)
        return wrapped
    return deco

BARCODE_DIR = os.path.join("static", "barcodes")
os.makedirs(BARCODE_DIR, exist_ok=True)

def get_db_connection():
    conn = sqlite3.connect('estoque.db')
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    with get_db_connection() as conn:
        cur = conn.cursor()

        # Itens de EPI
        cur.execute('''
            CREATE TABLE IF NOT EXISTS itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                codigo TEXT UNIQUE NOT NULL,
                saldo INTEGER NOT NULL
            )
        ''')

        # Histórico de saídas
        cur.execute('''
            CREATE TABLE IF NOT EXISTS movimentacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                quantidade INTEGER NOT NULL,
                destinatario TEXT NOT NULL,
                data TEXT NOT NULL,
                FOREIGN KEY (item_id) REFERENCES itens (id)
            )
        ''')

        # Fila de etiquetas a imprimir depois
        cur.execute('''
            CREATE TABLE IF NOT EXISTS etiquetas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                codigo TEXT NOT NULL,
                nome TEXT NOT NULL,
                numero_etiqueta INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'pendente',  -- 'pendente' | 'impresso'
                criado_em TEXT DEFAULT (datetime('now','localtime')),
                impresso_em TEXT,
                FOREIGN KEY (item_id) REFERENCES itens (id) ON DELETE CASCADE
            )
        ''')
        cur.execute("CREATE INDEX IF NOT EXISTS idx_etq_status ON etiquetas(status)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_etq_item   ON etiquetas(item_id)")

        conn.commit()
def ensure_ca_column(db_path='epi.db'):
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    try:
        # roda uma única vez; nas próximas dará OperationalError e ignoramos
        cur.execute("ALTER TABLE itens ADD COLUMN CA TEXT;")
        con.commit()
    except sqlite3.OperationalError:
        pass
    finally:
        con.close()

def proximo_numero_etiqueta(conn) -> int:
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(MAX(numero_etiqueta), 0) + 1 FROM etiquetas")
    return cur.fetchone()[0]

def gerar_codigo_se_vazio(conn, codigo_informado: str) -> str:
    """Se vazio, cria código interno sequencial EPI000001, EPI000002..."""
    cod = (codigo_informado or "").strip()
    if cod:
        return cod
    cur = conn.cursor()
    cur.execute("SELECT IFNULL(MAX(id), 0) + 1 AS prox FROM itens")
    prox = cur.fetchone()["prox"]
    return f"EPI{prox:06d}"

def salvar_barcode_png(
    codigo: str,
    *,
    module_width: float = 0.15,
    module_height: float = 8.0,
    font_size: int = 6,
    text_distance: float = 0.8,
    quiet_zone: float = 1.0,
    write_text: bool = False,   # <<< DESLIGA texto na imagem
    force: bool = False
) -> str:
    filename_sem_ext = os.path.join(BARCODE_DIR, codigo)
    path_png = f"{filename_sem_ext}.png"

    if not force and os.path.exists(path_png):
        return path_png

    try:
        if os.path.exists(path_png):
            os.remove(path_png)
    except Exception:
        pass

    writer = ImageWriter()
    options = {
        "module_width": module_width,
        "module_height": module_height,
        "font_size": font_size,
        "text_distance": text_distance,
        "quiet_zone": quiet_zone,
        "write_text": write_text,   # <<< aqui
    }

    b = Code39(codigo, writer=writer, add_checksum=False)
    b.save(filename_sem_ext, options=options)
    return path_png


def buscar_movimentacoes(destinatario=None, data_ini=None, data_fim=None):
    conn = get_db_connection()
    cur = conn.cursor()

    sql = """
      SELECT m.id, i.nome AS item_nome, i.codigo AS item_codigo,
             m.quantidade, m.destinatario, m.data
      FROM movimentacoes m
      JOIN itens i ON i.id = m.item_id
      WHERE 1=1
    """
    params = []

    if destinatario:
        sql += " AND LOWER(m.destinatario) = LOWER(?)"
        params.append(destinatario.strip())

    if data_ini:
        sql += " AND substr(m.data,1,10) >= ?"
        params.append(data_ini.strip())

    if data_fim:
        sql += " AND substr(m.data,1,10) <= ?"
        params.append(data_fim.strip())

    # 👉 agora mais novos primeiro
    sql += " ORDER BY m.data DESC, m.id DESC"

    cur.execute(sql, tuple(params))
    rows = cur.fetchall()
    conn.close()
    return rows


# --- util: converte 'DD/MM/AAAA' -> 'AAAA-MM-DD' (ou retorna '' se vazio/invalid) ---
def br_to_iso(d: str) -> str:
    d = (d or "").strip()
    if not d:
        return ""
    d = d.replace("-", "/")
    try:
        dd, mm, yy = d.split("/")
        if len(dd) == 2 and len(mm) == 2 and len(yy) == 4:
            return f"{yy}-{mm}-{dd}"
    except Exception:
        pass
    return ""  # se inválida, não aplica filtro

# -----------------Tela de BAIXA automática -----------------
@app.route("/", methods=["GET", "POST"])
@login_required_for(methods=("POST",))
def baixa_automatica():
    # helper para listar itens
    def listar_itens():
        conn = get_db_connection()
        cur = conn.cursor()
        # Garante que sempre venha CA e SALDO "seguros" ('' e 0 quando nulos)
        cur.execute("""
            SELECT 
                id,
                nome,
                codigo,
                COALESCE(ca, '')     AS ca,
                COALESCE(saldo, 0)    AS saldo
            FROM itens
            ORDER BY nome COLLATE NOCASE
        """)
        itens = cur.fetchall()
        conn.close()
        return itens

    # Detecta AJAX (para devolver JSON)
    wants_json = (
        request.headers.get("X-Requested-With") == "XMLHttpRequest"
        or "application/json" in (request.headers.get("Accept") or "")
    )

    if request.method == "POST":
        codigo = (request.form.get("codigo") or "").strip()
        destinatario = (request.form.get("destinatario") or "").strip()

        if not codigo:
            if wants_json:
                return jsonify({"ok": False, "erro": "Código vazio"}), 400
            return render_template("baixa.html", erro="Código vazio", itens=listar_itens())

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM itens WHERE codigo = ?", (codigo,))
        item = cur.fetchone()
        if not item:
            conn.close()
            if wants_json:
                return jsonify({"ok": False, "erro": f"Código {codigo} não encontrado."}), 404
            return render_template("baixa.html", erro=f"Código {codigo} não encontrado.", itens=listar_itens())

        # 🔒 Blindagem: se saldo vier NULL, trata como 0
        saldo_atual = item["saldo"] if item["saldo"] is not None else 0
        novo = max(0, saldo_atual - 1)

        cur.execute("UPDATE itens SET saldo = ? WHERE id = ?", (novo, item["id"]))
        cur.execute(
            "INSERT INTO movimentacoes (item_id, quantidade, destinatario, data) VALUES (?, ?, ?, ?)",
            (item["id"], 1, destinatario, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        conn.close()

        msg_ok = f"{item['nome']} (-1) para {destinatario or 'Sem nome'}. Saldo: {novo}"

        if wants_json:
            return jsonify({"ok": True, "restante": novo, "msg": msg_ok})

        return render_template("baixa.html", ok=True, msg=msg_ok, itens=listar_itens())

    # GET normal
    return render_template("baixa.html", itens=listar_itens())


#-------------login------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    # Se já estiver logado, manda para a página inicial
    if session.get("user"):
        return redirect(url_for("baixa_automatica"))

    erro = None
    if request.method == "POST":
        pwd = (request.form.get("senha") or "").strip()
        if pwd == ADMIN_PASS:  # usa a senha fixa já definida no topo do app
            session["user"] = "admin"  # usuário simbólico
            next_url = session.pop("next_url", None) or url_for("baixa_automatica")
            return redirect(next_url)
        erro = "Senha inválida."
    return render_template("login.html", erro=erro)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ----------------- Cadastrar NOVO item -----------------
@app.route("/novo", methods=["GET", "POST"])
@login_required_for(methods=("POST",))
def novo_item():
    """
    Cadastra um NOVO EPI:
    - Gera código se não informado.
    - Se código já existir, NÃO deixa salvar (evita duplicidade).
    - Gera etiqueta e oferece botão "Imprimir".
    """
    mensagem = None
    etiqueta_codigo = None

    if request.method == "POST":
        nome = request.form["nome"].strip()
        codigo_informado = (request.form.get("codigo") or "").strip()
        saldo_str = request.form["saldo"].strip()

        # valida saldo
        try:
            saldo = int(saldo_str)
            if saldo < 0:
                raise ValueError
        except ValueError:
            return render_template("novo.html", mensagem="Saldo deve ser inteiro ≥ 0.")

        conn = get_db_connection()
        cur = conn.cursor()

        # Se usuário informou um código, garantir que não exista
        if codigo_informado:
            cur.execute("SELECT 1 FROM itens WHERE codigo = ?", (codigo_informado,))
            if cur.fetchone():
                conn.close()
                return render_template("novo.html", mensagem=f"Código {codigo_informado} já existe. Use outro.")

        # Gera código (se vazio) com base no próximo id
        codigo = gerar_codigo_se_vazio(conn, codigo_informado)

        # Inserir novo item
        cur.execute("INSERT INTO itens (nome, codigo, saldo) VALUES (?, ?, ?)", (nome, codigo, saldo))
        conn.commit()
        conn.close()

        # Gera etiqueta
        salvar_barcode_png(codigo, force=True)
        mensagem = f"Item cadastrado: {nome} (cód. {codigo}), saldo {saldo}."
        etiqueta_codigo = codigo

        return render_template("novo.html", mensagem=mensagem, etiqueta_codigo=etiqueta_codigo)

    return render_template("novo.html", mensagem=mensagem, etiqueta_codigo=etiqueta_codigo)

# ----------------- Reposição em LISTA -----------------
@app.route("/repor", methods=["GET", "POST"])
@login_required_for(methods=("POST",))
def repor():
    """
    Reposição por LISTA com BUSCA:
    - GET: aceita parâmetro q (nome/código) e filtra a tabela.
    - POST:
        a) Se vier item_id e qtd -> atualiza apenas aquele item (botão por linha)
        b) Caso contrário -> modo antigo em lote (varre todos os campos qtd_<id>)
    """
    conn = get_db_connection()
    cur = conn.cursor()

    # ----- POST: MODO LINHA-A-LINHA -----
    if request.method == "POST" and request.form.get("item_id"):
        try:
            item_id = int(request.form.get("item_id"))
        except ValueError:
            item_id = 0
        try:
            qtd = int((request.form.get("qtd") or "0").strip())
        except ValueError:
            qtd = 0

        # busca item
        cur.execute("SELECT id, nome, saldo, codigo FROM itens WHERE id = ?", (item_id,))
        it = cur.fetchone()
        if not it:
            conn.close()
            # volta com erro simples
            q = (request.form.get("q") or "").strip()
            # recarrega lista
            conn2 = get_db_connection(); cur2 = conn2.cursor()
            if q:
                like = f"%{q}%"
                cur2.execute("""SELECT * FROM itens WHERE nome LIKE ? OR codigo LIKE ? ORDER BY nome""", (like, like))
            else:
                cur2.execute("SELECT * FROM itens ORDER BY nome")
            itens = cur2.fetchall(); conn2.close()
            return render_template("repor.html", itens=itens, q=q, resumo="Item não encontrado.")

        resumo = "Nenhuma quantidade informada."
        if qtd > 0:
            novo = it["saldo"] + qtd
            cur.execute("UPDATE itens SET saldo = ? WHERE id = ?", (novo, it["id"]))
            conn.commit()
            resumo = f"{it['nome']} +{qtd} (→ {novo})"

        # Após salvar 1 item, recarrega a lista (com o mesmo filtro q, se houver)
        q = (request.form.get("q") or "").strip()
        if q:
            like = f"%{q}%"
            cur.execute("""SELECT * FROM itens WHERE nome LIKE ? OR codigo LIKE ? ORDER BY nome""", (like, like))
        else:
            cur.execute("SELECT * FROM itens ORDER BY nome")
        itens = cur.fetchall()
        conn.close()
        return render_template("repor.html", itens=itens, q=q, resumo=resumo)

    # ----- POST: MODO EM LOTE (ANTIGO) -----
    if request.method == "POST":
        # Carrega todos para varrer os campos qtd_<id>
        cur.execute("SELECT id, nome, saldo FROM itens ORDER BY nome")
        itens_all = cur.fetchall()

        alterados = []
        for it in itens_all:
            field = f"qtd_{it['id']}"
            qtd_str = request.form.get(field, "0").strip()
            if not qtd_str:
                continue
            try:
                qtd = int(qtd_str)
            except ValueError:
                qtd = 0
            if qtd > 0:
                novo = it["saldo"] + qtd
                cur.execute("UPDATE itens SET saldo = ? WHERE id = ?", (novo, it["id"]))
                alterados.append((it["nome"], qtd, novo))

        if alterados:
            conn.commit()

        conn.close()
        resumo = ", ".join([f"{nome} +{qtd} (→ {novo})" for (nome, qtd, novo) in alterados]) or "Nenhuma quantidade informada."
        return render_template("repor_resultado.html", resumo=resumo)

    # ----- GET → busca e lista -----
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        cur.execute("""
            SELECT * FROM itens
            WHERE nome LIKE ? OR codigo LIKE ?
            ORDER BY nome
        """, (like, like))
    else:
        cur.execute("SELECT * FROM itens ORDER BY nome")

    itens = cur.fetchall()
    conn.close()
    return render_template("repor.html", itens=itens, q=q)


# ----------------- LISTA DE ITENS (gerenciar) -----------------
@app.route("/itens")
@login_required_for(methods=("POST",))
def itens_lista():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM itens ORDER BY nome")
    itens = cur.fetchall()
    conn.close()
    return render_template("itens.html", itens=itens)

# ----------------- EDITAR ITEM -----------------
@app.route("/editar/<int:item_id>", methods=["GET", "POST"])
@login_required
def editar_item(item_id):
    """
    Edita nome, código e saldo manualmente.
    - Garante código único (pode manter o mesmo).
    - Se o código mudar, atualiza a etiqueta (gera PNG novo e remove PNG antigo).
    """
    conn = get_db_connection()
    cur = conn.cursor()

    # Busca o item atual
    cur.execute("SELECT * FROM itens WHERE id = ?", (item_id,))
    item = cur.fetchone()
    if not item:
        conn.close()
        return render_template("editar.html", erro="Item não encontrado.", item=None)

    if request.method == "POST":
        nome = (request.form.get("nome") or "").strip()
        codigo = (request.form.get("codigo") or "").strip()
        saldo_str = (request.form.get("saldo") or "").strip()

        # Validações
        if not nome:
            conn.close()
            return render_template("editar.html", erro="Nome é obrigatório.", item=item)

        if not codigo:
            conn.close()
            return render_template("editar.html", erro="Código é obrigatório.", item=item)

        try:
            saldo = int(saldo_str)
            if saldo < 0:
                raise ValueError
        except ValueError:
            conn.close()
            return render_template("editar.html", erro="Saldo deve ser inteiro ≥ 0.", item=item)

        # Código único (permitindo o mesmo do próprio item)
        cur.execute("SELECT id FROM itens WHERE codigo = ? AND id <> ?", (codigo, item_id))
        conflito = cur.fetchone()
        if conflito:
            conn.close()
            return render_template("editar.html", erro=f"Código {codigo} já existe em outro item.", item=item)

        codigo_antigo = item["codigo"]

        # Atualiza
        cur.execute("UPDATE itens SET nome = ?, codigo = ?, saldo = ? WHERE id = ?",
                    (nome, codigo, saldo, item_id))
        conn.commit()
        conn.close()

        # Se código mudou, atualiza etiqueta: apaga PNG antigo e gera novo
        if codigo != codigo_antigo:
            antigo_png = os.path.join(BARCODE_DIR, f"{codigo_antigo}.png")
            if os.path.exists(antigo_png):
                try:
                    os.remove(antigo_png)
                except Exception:
                    pass
            salvar_barcode_png(codigo, force=True)

        return redirect(url_for("itens_lista"))

    # GET
    conn.close()
    return render_template("editar.html", item=item, erro=None)

# ----------------- EXCLUIR ITEM -----------------
@app.route("/excluir/<int:item_id>", methods=["GET", "POST"])
@login_required
def excluir_item(item_id):
    """
    Exclui item com confirmação.
    - Remove também todas as movimentações vinculadas (ON DELETE CASCADE).
    """
    conn = get_db_connection()
    cur = conn.cursor()

    # Busca o item
    cur.execute("SELECT * FROM itens WHERE id = ?", (item_id,))
    item = cur.fetchone()
    if not item:
        conn.close()
        return render_template("excluir.html", item=None, erro="Item não encontrado.")

    if request.method == "POST":
        # Remove PNG da etiqueta (se existir)
        png_path = os.path.join(BARCODE_DIR, f"{item['codigo']}.png")
        if os.path.exists(png_path):
            try:
                os.remove(png_path)
            except Exception:
                pass

        # Excluir movimentações vinculadas
        cur.execute("DELETE FROM movimentacoes WHERE item_id = ?", (item_id,))
        # Excluir item
        cur.execute("DELETE FROM itens WHERE id = ?", (item_id,))
        conn.commit()
        conn.close()
        return redirect(url_for("itens_lista"))

    conn.close()
    return render_template("excluir.html", item=item, erro=None)

# ----------------- RELATORIOS-----------------
@app.route("/relatorios", methods=["GET"])
def relatorios():
    # leitura do formulário
    destinatario = (request.args.get("destinatario") or "").strip()
    data_ini_br  = (request.args.get("data_ini") or "").strip()
    data_fim_br  = (request.args.get("data_fim") or "").strip()

    # conversão para ISO
    data_ini_iso = br_to_iso(data_ini_br)
    data_fim_iso = br_to_iso(data_fim_br)

    # 👉 sempre busca: se vier tudo vazio, retorna TODAS as movimentações
    movimentos = buscar_movimentacoes(
        destinatario if destinatario else None,
        data_ini_iso if data_ini_iso else None,
        data_fim_iso if data_fim_iso else None
    )

    # sugestões de destinatários
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT destinatario FROM movimentacoes ORDER BY destinatario")
    destinatarios_unicos = [r["destinatario"] for r in cur.fetchall()]
    conn.close()

    return render_template(
        "relatorios.html",
        movimentos=movimentos,
        destinatario=destinatario,
        data_ini=data_ini_br,
        data_fim=data_fim_br,
        destinatarios_unicos=destinatarios_unicos
    )

# ----------------- RELATORIOS/EXPORT-----------------
@app.route("/relatorios/export", methods=["GET"])
def relatorios_export():
    destinatario = (request.args.get("destinatario") or "").strip()
    data_ini_iso = br_to_iso(request.args.get("data_ini"))
    data_fim_iso = br_to_iso(request.args.get("data_fim"))

    movimentos = buscar_movimentacoes(destinatario, data_ini_iso, data_fim_iso)

    wb = Workbook()
    ws = wb.active
    ws.title = "Envios EPI"
    ws.append(["Destinatário", "Item", "Código", "Quantidade", "Data/Hora"])

    for mv in movimentos:
        # mv["data"] é 'AAAA-MM-DD HH:MM:SS' -> exibir BR no Excel
        data_br = f"{mv['data'][8:10]}/{mv['data'][5:7]}/{mv['data'][0:4]} {mv['data'][11:]}"
        ws.append([mv["destinatario"], mv["item_nome"], mv["item_codigo"], mv["quantidade"], data_br])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name="relatorio_envios.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ----------------- Etiqueta para impressão -----------------
@app.route("/etiqueta/<codigo>")
def etiqueta(codigo):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT nome FROM itens WHERE codigo = ?", (codigo,))
    item = cur.fetchone()
    conn.close()

    if not item:
        nome = "(Item não encontrado)"
    else:
        nome = item["nome"]

    salvar_barcode_png(codigo)  # garante que o PNG exista
    return render_template("etiqueta.html", codigo=codigo, nome=nome)

# --- Escolher/guardar destinatário no cookie (uma vez) ---
@app.route("/destinatario", methods=["GET", "POST"])
def escolher_destinatario():
    """
    GET: mostra um formulário simples para definir o destinatário.
    POST: salva o destinatário em cookie e volta para a própria página (ok=1).
    """
    if request.method == "POST":
        dest = (request.form.get("destinatario") or "").strip()
        if not dest:
            return render_template("set_destinatario.html", atual="", ok=False, erro="Informe um nome.")
        resp = make_response(redirect(url_for("escolher_destinatario", ok=1)))
        # cookie por 30 dias
        resp.set_cookie("destinatario", dest, max_age=60*60*24*30)
        return resp

    atual = request.cookies.get("destinatario", "")
    ok = request.args.get("ok") == "1"
    return render_template("set_destinatario.html", atual=atual, ok=ok, erro=None)

@app.route("/ping")
def ping():
    return "pong", 200

# -------- Enfileirar etiqueta (para imprimir depois) ----------
@app.route("/etiquetas/enfileirar", methods=["POST"])
def etiquetas_enfileirar():
    """
    Body JSON: { "item_id": <int> } OU { "codigo": "<str>" }.
    Gera uma etiqueta 'pendente' com número sequencial.
    """
    data = request.get_json(force=True) or {}
    item_id = data.get("item_id")
    codigo_in = (data.get("codigo") or "").strip()

    conn = get_db_connection(); cur = conn.cursor()
    try:
        if item_id:
            cur.execute("SELECT id, nome, codigo FROM itens WHERE id = ?", (int(item_id),))
        else:
            cur.execute("SELECT id, nome, codigo FROM itens WHERE codigo = ?", (codigo_in,))
        row = cur.fetchone()
        if not row:
            conn.close()
            return {"ok": False, "msg": "Item não encontrado."}, 404

        iid, nome, codigo = row["id"], row["nome"], row["codigo"]
        # Número sequencial
        numero = proximo_numero_etiqueta(conn)

        # Snapshot (nome e codigo gravados na etiqueta)
        cur.execute("""
            INSERT INTO etiquetas (item_id, codigo, nome, numero_etiqueta, status)
            VALUES (?, ?, ?, ?, 'pendente')
        """, (iid, codigo, nome, numero))
        conn.commit()

        # Garante que o PNG existe
        salvar_barcode_png(codigo, force=True)

        return {"ok": True, "id": cur.lastrowid, "numero_etiqueta": numero, "codigo": codigo, "nome": nome}
    except Exception as e:
        conn.rollback()
        return {"ok": False, "msg": f"Erro: {e}"}, 500
    finally:
        conn.close()

# -------- Lista de etiquetas pendentes ----------
@app.route("/etiquetas/pendentes")
def etiquetas_pendentes():
    conn = get_db_connection(); cur = conn.cursor()
    cur.execute("""
      SELECT id, numero_etiqueta, nome, codigo, status
      FROM etiquetas
      WHERE status='pendente'
      ORDER BY id DESC
    """)
    rows = cur.fetchall()
    conn.close()
    return render_template("etiquetas_pendentes.html", rows=rows)

# -------- Página de impressão (selecionadas ou todas pendentes) ----------
@app.route("/etiquetas/print")
def etiquetas_print():
    ids_raw = (request.args.get("ids") or "").strip()   # ex: "5,6,7"
    conn = get_db_connection(); cur = conn.cursor()

    if ids_raw:
        idlist = [int(x) for x in ids_raw.split(",") if x.strip().isdigit()]
        if idlist:
            qmarks = ",".join(["?"]*len(idlist))
            cur.execute(f"""
                SELECT id, numero_etiqueta, nome, codigo
                FROM etiquetas WHERE id IN ({qmarks})
            """, tuple(idlist))
        else:
            cur.execute("SELECT id, numero_etiqueta, nome, codigo FROM etiquetas WHERE 1=0")
    else:
        cur.execute("""
            SELECT id, numero_etiqueta, nome, codigo
            FROM etiquetas WHERE status='pendente'
        """)

    etiquetas = cur.fetchall()
    conn.close()

    # Garante PNG de todas antes de renderizar
    for e in etiquetas:
        salvar_barcode_png(e["codigo"], force=True)

    return render_template("etiquetas_print.html", etiquetas=etiquetas)

# -------- Marcar selecionadas como impressas ----------
@app.route("/etiquetas/marcar_impresso", methods=["POST"])
def etiquetas_marcar_impresso():
    data = request.get_json(force=True) or {}
    ids = data.get("ids") or []  # lista de ints

    if not ids:
        return {"ok": False, "msg": "Sem IDs."}, 400

    conn = get_db_connection(); cur = conn.cursor()
    try:
        qmarks = ",".join(["?"]*len(ids))
        cur.execute(f"""
          UPDATE etiquetas
             SET status='impresso', impresso_em=datetime('now','localtime')
           WHERE id IN ({qmarks})
        """, tuple(ids))
        conn.commit()
        return {"ok": True, "atualizadas": cur.rowcount}
    except Exception as e:
        conn.rollback()
        return {"ok": False, "msg": f"Erro: {e}"}, 500
    finally:
        conn.close()

#-------------reimprimir---------
@app.route("/etiquetas/historico")
def etiquetas_historico():
    conn = get_db_connection(); cur = conn.cursor()
    cur.execute("""
      SELECT id, item_id, numero_etiqueta, nome, codigo, status, criado_em, impresso_em
      FROM etiquetas
      ORDER BY id DESC
    """)
    rows = cur.fetchall()
    conn.close()
    return render_template("etiquetas_historico.html", rows=rows)

#-------------EXPORTAR ESTOQUE ATUAL--------------
@app.route("/estoque/export", methods=["GET"])
def estoque_export():
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("""
        SELECT 
            nome,
            COALESCE(ca, '')   AS ca,
            codigo,
            COALESCE(saldo, 0) AS saldo
        FROM itens
        ORDER BY nome COLLATE NOCASE
    """)
    rows = cur.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Estoque"
    # Cabeçalho na ordem pedida
    ws.append(["Item", "CA", "Código", "Saldo"])

    for r in rows:
        ws.append([r["nome"], r["ca"], r["codigo"], r["saldo"]])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    nome_arquivo = f"estoque_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=False)




#  git add .
#  git commit -m "Relatorios de estoque e tabela de CA"
#  git push
