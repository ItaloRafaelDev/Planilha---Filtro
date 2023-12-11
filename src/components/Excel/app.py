from flask import Flask, request, send_file
import pandas as pd

app = Flask(__name__)

@app.route('/gerar_planilha', methods=['POST'])
def gerar_planilha():
    dados = request.json['dados']  # Receba os dados do frontend

    # Processamento dos dados com Pandas e OpenPyXL
    df = pd.DataFrame(dados)
    # Realize operações com os dados, adicione filtros, formatações, etc.

    # Crie a planilha usando Pandas e salve o arquivo localmente ou em memória
    df.to_excel('planilha_gerada.xlsx', index=False)

    # Envie a planilha de volta para o frontend
    return send_file('planilha_gerada.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(host='localhost', port=5000)  # Execute o servidor na máquina local, porta 5000
