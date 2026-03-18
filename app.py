# app.py
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import uuid
import io
import os
import tempfile
from pathlib import Path

import parser_core

parse_docx_questions = parser_core.parse_docx_questions


def _resolve_xml_exporter(module):
    """
    Resolve compativelmente o exportador de XML entre versões do parser_core.
    """
    candidates = [
        "make_moodle_xml",
        "build_moodle_xml",
        "export_moodle_xml",
        "to_moodle_xml",
    ]
    for name in candidates:
        fn = getattr(module, name, None)
        if callable(fn):
            return fn
    return None


make_moodle_xml = _resolve_xml_exporter(parser_core)

if make_moodle_xml is None:
    available = sorted(
        name
        for name in dir(parser_core)
        if name.endswith("xml") or name.endswith("moodle_xml")
    )
    raise ImportError(
        "parser_core não expõe uma função de exportação XML. "
        "Esperado: make_moodle_xml (ou aliases compatíveis). "
        f"Encontrado: {', '.join(available) if available else '(nenhuma função XML encontrada)'}"
    )

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB
app.config["UPLOAD_EXTENSIONS"] = [".docx"]

# 📌 Diretório temporário seguro (cross-platform)
TMP_DIR = Path(tempfile.gettempdir()) / "docx2moodle"
TMP_DIR.mkdir(parents=True, exist_ok=True)

# memstore simples: id -> questões separadas por visibilidade
MEM = {}

@app.route("/")
def index():
    return render_template("index.html")

@app.post("/api/parse")
def api_parse():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Nenhum arquivo enviado."}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"ok": False, "error": "Arquivo sem nome."}), 400

    fn = secure_filename(f.filename)
    ext = os.path.splitext(fn)[1].lower()
    if ext not in app.config["UPLOAD_EXTENSIONS"]:
        return jsonify({"ok": False, "error": "Apenas .docx é aceito."}), 400

    # 📥 salva temporariamente no diretório certo
    tmp_path = TMP_DIR / f"{uuid.uuid4()}.docx"
    f.save(str(tmp_path))

    try:
        questoes = parse_docx_questions(str(tmp_path))
    except Exception as e:
        return jsonify({"ok": False, "error": f"Falha ao ler DOCX: {e}"}), 500
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except:
            pass

    if not questoes:
        return jsonify({"ok": False, "error": "Nenhuma questão encontrada. Confirme o padrão do DOCX."}), 200

    questoes_visiveis = [q for q in questoes if not q.get("discarded")]
    questoes_descartadas = [q for q in questoes if q.get("discarded")]

    doc_id = str(uuid.uuid4())
    MEM[doc_id] = {
        "all": questoes,
        "visible": questoes_visiveis,
        "discarded": questoes_descartadas,
    }

    return jsonify({
        "ok": True,
        "doc_id": doc_id,
        "count": len(questoes_visiveis),
        "discarded_count": len(questoes_descartadas),
        "questoes": questoes_visiveis,
    })

@app.post("/api/export")
def api_export():
    data = request.get_json(silent=True) or {}
    doc_id = data.get("doc_id")
    marks = data.get("mark_remove", [])  # lista de "U#Q#"
    categoria = data.get("categoria", " Metodologia do Ensino de Ciências")

    if not doc_id or doc_id not in MEM:
        return jsonify({"ok": False, "error": "doc_id inválido ou expirado."}), 400

    # 🏷️ converte "U1Q2" → (1,2)
    mark_set = set()
    for item in marks:
        item = item.strip().upper()
        if item.startswith("U") and "Q" in item:
            try:
                u = int(item.split("Q")[0][1:])
                q = int(item.split("Q")[1])
                mark_set.add((u, q))
            except:
                pass

    store = MEM[doc_id]
    questoes = store["all"]
    tree = make_moodle_xml(questoes, mark_remove_set=mark_set, categoria=categoria)

    # 📤 retorna como arquivo XML para download
    xml_bytes = io.BytesIO()
    tree.write(xml_bytes, encoding="utf-8", xml_declaration=True, pretty_print=True)
    xml_bytes.seek(0)

    filename = f"moodle_{doc_id}.xml"
    return send_file(xml_bytes, as_attachment=True, download_name=filename, mimetype="application/xml")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
