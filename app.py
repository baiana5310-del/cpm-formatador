from flask import Flask, render_template, request, jsonify, send_file
from pathlib import Path
from werkzeug.utils import secure_filename
from engine import processar_arquivos

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/processar", methods=["POST"])
def processar():
    arquivos = request.files.getlist("arquivos")
    tema = request.form.get("tema", "🔵 Azul Executivo")

    if not arquivos or arquivos[0].filename == "":
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400

    caminhos = []

    for arquivo in arquivos:
        nome = secure_filename(arquivo.filename)
        caminho = UPLOAD_DIR / nome
        arquivo.save(caminho)
        caminhos.append(str(caminho))

    try:
        print("📥 Arquivos recebidos:", caminhos)
        print("🎨 Tema recebido:", tema)

        saida = processar_arquivos(
            caminhos,
            str(OUTPUT_DIR),
            tema=tema
        )
        saida_path = Path(saida)

        print("📤 Arquivo gerado:", saida_path)

        if not saida_path.exists():
            return jsonify({"erro": f"Arquivo não encontrado: {saida_path}"}), 500

        return jsonify({
            "ok": True,
            "download": f"/download/{saida_path.name}"
        })

    except Exception as e:
        print("❌ ERRO NO PROCESSAMENTO:", e)
        return jsonify({"erro": str(e)}), 500

@app.route("/download/<nome>")
def download(nome):
    caminho = OUTPUT_DIR / nome
    print("⬇️ Download solicitado:", caminho)

    if not caminho.exists():
        return f"Arquivo não encontrado: {caminho}", 404

    return send_file(caminho, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)