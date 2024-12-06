import os
import io
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import base64
from flask import Flask, request, render_template, send_file
from openpyxl import load_workbook
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
    file_extension = os.path.splitext(file_path)[1]
    if file_extension == '.xls':
        df = pd.read_excel(file_path, engine='xlrd')
    else:
        df = pd.read_excel(file_path, engine='openpyxl')

    expected_columns = {
        'ID do ticket': 'ID do ticket',
        'Hora da resolução': 'Hora da resolução',
        'Primeiro prazo': 'Primeiro prazo',
        'Nome completo': 'Nome completo'
    }
    for original, corrected in expected_columns.items():
        if original in df.columns:
            df.rename(columns={original: corrected}, inplace=True)

    missing_columns = [col for col in expected_columns.values() if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Colunas esperadas não encontradas: {', '.join(missing_columns)}")

    df['Hora da resolução'] = pd.to_datetime(df['Hora da resolução'], errors='coerce').dt.date
    df['Primeiro prazo'] = pd.to_datetime(df['Primeiro prazo'], dayfirst=True, errors='coerce')
    df['Dias de diferença'] = (pd.to_datetime(df['Hora da resolução']) - df['Primeiro prazo']).dt.days

    df['Status'] = df.apply(
        lambda row: 'Sem prazo' if pd.isna(row['Primeiro prazo']) else (
            'Fora do prazo' if row['Dias de diferença'] > 0 else 'No prazo'
        ),
        axis=1
    )

    total_tickets = len(df)
    no_prazo = len(df[df['Status'] == 'No prazo'])
    fora_prazo = len(df[df['Status'] == 'Fora do prazo'])
    sem_prazo = len(df[df['Status'] == 'Sem prazo'])

    percent_no_prazo = (no_prazo / total_tickets) * 100
    percent_fora_prazo = (fora_prazo / total_tickets) * 100
    percent_sem_prazo = (sem_prazo / total_tickets) * 100

    labels = ['No Prazo', 'Fora do Prazo', 'Sem Prazo']
    sizes = [no_prazo, fora_prazo, sem_prazo]
    colors = ['#28a745', '#dc3545', '#6c757d']
    explode = (0.1, 0.1, 0)

    plt.figure(figsize=(6, 6))
    plt.pie(sizes, labels=labels, colors=colors, explode=explode, autopct='%1.1f%%', shadow=True, startangle=140)
    plt.axis('equal')

    buffer = io.BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    graph_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
    buffer.close()

    ranking = df['Nome completo'].value_counts().reset_index()
    ranking.columns = ['Cliente', 'Quantidade de Solicitações']
    ranking = ranking.sort_values(by='Quantidade de Solicitações', ascending=False).head(10)

    output_file = os.path.join(UPLOAD_FOLDER, 'output.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Dados Processados', index=False)
        pd.DataFrame({
            'Status': ['No prazo', 'Fora do prazo', 'Sem prazo'],
            'Quantidade': [no_prazo, fora_prazo, sem_prazo],
            'Porcentagem': [percent_no_prazo, percent_fora_prazo, percent_sem_prazo]
        }).to_excel(writer, sheet_name='Resumo', index=False)
        ranking.to_excel(writer, sheet_name='Ranking', index=False)

    global tickets_fora_prazo
    tickets_fora_prazo = df[df['Status'] == 'Fora do prazo'][['ID do ticket', 'Primeiro prazo', 'Hora da resolução']]
    return df, output_file, graph_data, no_prazo, fora_prazo, sem_prazo, ranking

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files or request.files['file'].filename == '':
            return 'Nenhum arquivo enviado.'

        clear_upload_folder()
        file = request.files['file']
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        df, output_file, graph_data, no_prazo, fora_prazo, sem_prazo, ranking = process_excel(file_path)
        return render_template('result.html', df=df.to_html(classes="table table-bordered"), graph_data=graph_data,
                               no_prazo=no_prazo, fora_prazo=fora_prazo, sem_prazo=sem_prazo,
                               ranking=ranking.to_html(classes="table table-bordered"))

    return render_template('upload.html')

@app.route('/fora_do_prazo')
def fora_do_prazo():
    if tickets_fora_prazo.empty:
        print("Não há tickets fora do prazo.")  # Log de depuração
        return render_template('fora_do_prazo.html', tickets=[])

    # Calcular os dias de diferença para cada ticket
    tickets_fora_prazo['Dias de diferença'] = (tickets_fora_prazo['Hora da resolução'] - tickets_fora_prazo['Primeiro prazo']).dt.days

    # Criar link para o ticket
    tickets_fora_prazo['Link'] = tickets_fora_prazo['ID do ticket'].apply(
        lambda x: f'<a href="https://atendimento.p21sistemas.com.br/a/tickets/{x}" target="_blank">Ticket {x}</a>'
    )

    # Retornar para a página com os tickets fora do prazo
    return render_template('fora_do_prazo.html', tickets=tickets_fora_prazo.to_dict(orient='records'))

@app.route('/download')
def download_file():
    return send_file(os.path.join(UPLOAD_FOLDER, 'output.xlsx'), as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(host='0.0.0.0', port=5000, debug=True)
