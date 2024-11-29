import os
import io
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import base64
from flask import Flask, request, render_template, send_file
from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from werkzeug.utils import secure_filename

# Configurar Matplotlib para não usar interface gráfica
matplotlib.use('Agg')

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Função para limpar arquivos antigos antes de salvar um novo
def clear_upload_folder():
    for file in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

# Função para processar a planilha
def process_excel(file_path):
    # Verificar a extensão do arquivo
    file_extension = os.path.splitext(file_path)[1]
    if file_extension == '.xls':
        df = pd.read_excel(file_path, engine='xlrd')
    else:
        df = pd.read_excel(file_path, engine='openpyxl')

    # Verificar as colunas do DataFrame
    print("Colunas do DataFrame:", df.columns)

    # Ajustar os nomes das colunas se necessário
    expected_columns = {
        'ID do ticket': 'ID do ticket',
        'Hora da resolução': 'Hora da resolução',
        'Primeiro prazo': 'Primeiro prazo'
    }

    for original, corrected in expected_columns.items():
        if original in df.columns:
            df.rename(columns={original: corrected}, inplace=True)
        else:
            print(f"Coluna '{original}' não encontrada no DataFrame.")

    # Verificar se todas as colunas esperadas estão presentes
    missing_columns = [col for col in expected_columns.values() if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Colunas esperadas não encontradas: {', '.join(missing_columns)}")

    # Tratar as colunas de datas
    df['Hora da resolução'] = pd.to_datetime(df['Hora da resolução'], errors='coerce').dt.date
    df['Primeiro prazo'] = pd.to_datetime(df['Primeiro prazo'], dayfirst=True, errors='coerce')

    # Verificar se há datas inválidas
    invalid_dates = df[df['Primeiro prazo'].isna()]
    if not invalid_dates.empty:
        print(f"Datas inválidas encontradas na coluna 'Primeiro prazo':\n{invalid_dates[['ID do ticket', 'Primeiro prazo']]}")

    # Calcular a diferença entre o prazo e a resolução
    df['Dias de diferença'] = (pd.to_datetime(df['Hora da resolução']) - df['Primeiro prazo']).dt.days

    # Ajustar status (No prazo, Fora do prazo, Sem prazo)
    df['Status'] = df.apply(
        lambda row: 'Sem prazo' if pd.isna(row['Primeiro prazo']) else (
            'Fora do prazo' if row['Dias de diferença'] > 0 else 'No prazo'
        ),
        axis=1
    )

    # Contar quantidades e porcentagens
    total_tickets = len(df)
    no_prazo = len(df[df['Status'] == 'No prazo'])
    fora_prazo = len(df[df['Status'] == 'Fora do prazo'])
    sem_prazo = len(df[df['Status'] == 'Sem prazo'])

    # Evitar divisão por zero
    if total_tickets > 0:
        no_prazo_percent = (no_prazo / total_tickets) * 100
        fora_prazo_percent = (fora_prazo / total_tickets) * 100
        sem_prazo_percent = (sem_prazo / total_tickets) * 100
    else:
        no_prazo_percent = fora_prazo_percent = sem_prazo_percent = 0

    # Criar um DataFrame para armazenar as porcentagens e quantidades
    porcentagens_df = pd.DataFrame({
        'Status': ['No prazo', 'Fora do prazo', 'Sem prazo', 'Total'],
        'Porcentagem (%)': [no_prazo_percent, fora_prazo_percent, sem_prazo_percent, 100],
        'Quantidade': [no_prazo, fora_prazo, sem_prazo, total_tickets]
    })

    # Gerar gráfico de barras com Matplotlib
    fig, ax = plt.subplots()
    statuses = ['No prazo', 'Fora do prazo', 'Sem prazo']
    percentages = [no_prazo_percent, fora_prazo_percent, sem_prazo_percent]
    ax.bar(statuses, percentages, color=['green', 'red', 'gray'])
    ax.set_ylabel('Porcentagem (%)')
    ax.set_title('Distribuição Percentual dos Tickets')

    # Salvar o gráfico em um buffer de memória e convertê-lo para uma string base64
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    graph_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
    plt.close(fig)

    # Salvar a planilha processada em uma aba e as porcentagens em outra aba
    output_file = os.path.join(UPLOAD_FOLDER, 'output.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Dados Processados', index=False)
        porcentagens_df.to_excel(writer, sheet_name='Porcentagens', index=False)

    # Adicionar gráfico de pizza diretamente no Excel
    wb = load_workbook(output_file)
    ws = wb['Porcentagens']

    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=2, max_row=4)
    data = Reference(ws, min_col=3, min_row=2, max_row=4)
    pie.add_data(data, titles_from_data=False)
    pie.set_categories(labels)
    pie.title = "Distribuição de Tickets"

    ws.add_chart(pie, "E5")
    wb.save(output_file)

    return df, output_file, graph_data

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'Nenhum arquivo enviado'

        file = request.files['file']
        if file.filename == '':
            return 'Nenhum arquivo selecionado'

        # Limpar arquivos antigos e salvar o novo arquivo
        clear_upload_folder()
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        try:
            df, output_file, graph_data = process_excel(file_path)
            return render_template('result.html', graph_data=graph_data)
        except Exception as e:
            return f"Ocorreu um erro: {str(e)}"

    return render_template('upload.html')

@app.route('/download')
def download_file():
    return send_file(os.path.join(UPLOAD_FOLDER, 'output.xlsx'), as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(host='0.0.0.0', port=5000, debug=True)
