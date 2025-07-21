import os
import io
import pandas as pd
import numpy as np # Adicionado para a lógica de qualidade
import matplotlib
import matplotlib.pyplot as plt
import base64
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from dateutil import parser

# Configuração do Matplotlib para funcionar sem interface gráfica
matplotlib.use('Agg')

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Variável global para tickets fora do prazo (já existia)
tickets_fora_prazo = pd.DataFrame()

# Nomes das colunas para a análise de qualidade
COLUNAS_QUALIDADE = [
    'Enunciado claro?',
    'Alinhado prazo?',
    'Registrou o atendimento?',
    'Resposta clara para o cliente?',
    'Atendido no prazo?',
    'Manteve o cliente atualizado?'
]


def clear_upload_folder():
    """Limpa a pasta de uploads antes de um novo envio."""
    for file in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, file)
        if os.path.isfile(file_path):
            os.remove(file_path)


def process_excel(file_path):
    """
    Função principal que processa o arquivo enviado, realizando a análise
    de prazos e a nova análise de qualidade.
    """
    global tickets_fora_prazo

    # --- 1. LEITURA DO ARQUIVO (Lógica existente) ---
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.csv':
        try:
            df = pd.read_csv(file_path, encoding='utf-8', sep=',')
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding='latin1', sep=',')
    elif file_extension in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Formato de arquivo não suportado.")

    # --- 2. ANÁLISE DE PRAZOS (Lógica existente) ---
    required_cols_prazo = ['ID do ticket', 'Hora da resolução', 'Primeiro prazo', 'Nome completo']
    if not all(col in df.columns for col in required_cols_prazo):
        raise ValueError(f"Para análise de prazo, as colunas esperadas não foram encontradas: {', '.join(required_cols_prazo)}")
        
    df['Hora da resolução'] = pd.to_datetime(df['Hora da resolução'], errors='coerce')
    df['Primeiro prazo'] = df['Primeiro prazo'].apply(lambda x: parser.parse(x, dayfirst=True) if pd.notna(x) else pd.NaT)
    df['Dias de diferença'] = (df['Hora da resolução'] - df['Primeiro prazo']).dt.days
    df['Status Prazo'] = df.apply(
        lambda row: 'Sem prazo' if pd.isna(row['Primeiro prazo']) or pd.isna(row['Hora da resolução']) else ('Fora do prazo' if row['Dias de diferença'] > 0 else 'No prazo'),
        axis=1
    )
    tickets_fora_prazo = df[df['Status Prazo'] == 'Fora do prazo'][['ID do ticket', 'Primeiro prazo', 'Hora da resolução']]
    
    # Resumo da análise de prazo
    total_tickets = len(df)
    no_prazo = len(df[df['Status Prazo'] == 'No prazo'])
    fora_prazo = len(df[df['Status Prazo'] == 'Fora do prazo'])
    sem_prazo = len(df[df['Status Prazo'] == 'Sem prazo'])
    
    # Gráfico de pizza para prazos
    plt.figure(figsize=(6, 6))
    plt.pie([no_prazo, fora_prazo, sem_prazo], labels=['No Prazo', 'Fora do Prazo', 'Sem Prazo'], 
            colors=['#28a745', '#dc3545', '#6c757d'], explode=(0.1, 0.1, 0), autopct='%1.1f%%', shadow=True, startangle=140)
    plt.axis('equal')
    buffer_prazo = io.BytesIO()
    plt.savefig(buffer_prazo, format='png')
    plt.close()
    buffer_prazo.seek(0)
    graph_data_prazo = base64.b64encode(buffer_prazo.getvalue()).decode('utf-8')

    # Ranking de clientes
    ranking_clientes = df['Nome completo'].value_counts().reset_index()
    ranking_clientes.columns = ['Cliente', 'Quantidade de Solicitações']
    ranking_clientes = ranking_clientes.head(10)

    # --- 3. NOVA ANÁLISE DE QUALIDADE (Lógica integrada) ---
    qualidade_results = {}
    # Verifica se as colunas de qualidade existem no arquivo
    if all(col in df.columns for col in COLUNAS_QUALIDADE):
        # Adiciona a coluna 'Qualidade' (Aprovado/Reprovado)
        condicao_reprovado = (df[COLUNAS_QUALIDADE] == 'Não').any(axis=1)
        df['Qualidade Geral'] = np.where(condicao_reprovado, 'Reprovado', 'Aprovado')

        # Cria o DataFrame de resumo da qualidade
        resumo_metricas = {}
        for coluna in COLUNAS_QUALIDADE:
            resumo_metricas[coluna] = (df[coluna] == 'Sim').mean()
        resumo_metricas['**QUALIDADE GERAL**'] = (df['Qualidade Geral'] == 'Aprovado').mean()
        
        df_qualidade_resumo = pd.DataFrame(resumo_metricas.items(), columns=['Métrica', 'Percentual'])
        
        # Gráfico de barras para qualidade
        plt.figure(figsize=(10, 6))
        # Exclui o total para plotar apenas as perguntas
        plot_data = df_qualidade_resumo[~df_qualidade_resumo['Métrica'].str.contains('QUALIDADE GERAL')]
        bars = plt.barh(plot_data['Métrica'], plot_data['Percentual'] * 100, color='skyblue')
        plt.xlabel('Percentual de Conformidade (%)')
        plt.title('Análise de Qualidade por Pergunta')
        plt.xlim(0, 100)
        plt.gca().invert_yaxis() # Pergunta de cima para baixo
        # Adicionar os valores nas barras
        for bar in bars:
            plt.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, f'{bar.get_width():.1f}%', va='center')

        buffer_qualidade = io.BytesIO()
        plt.savefig(buffer_qualidade, format='png', bbox_inches='tight')
        plt.close()
        buffer_qualidade.seek(0)
        graph_data_qualidade = base64.b64encode(buffer_qualidade.getvalue()).decode('utf-8')
        
        # Formata o resumo para exibição
        df_qualidade_resumo['Percentual'] = df_qualidade_resumo['Percentual'].map('{:.2%}'.format)
        qualidade_results = {
            "resumo_df": df_qualidade_resumo,
            "graph_data": graph_data_qualidade
        }
    
    # --- 4. SALVAR ARQUIVO EXCEL COM TODAS AS ABAS ---
    output_file = os.path.join(UPLOAD_FOLDER, 'output.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Dados Processados', index=False)
        pd.DataFrame({
            'Status': ['No prazo', 'Fora do prazo', 'Sem prazo'],
            'Quantidade': [no_prazo, fora_prazo, sem_prazo],
            'Porcentagem': [(no_prazo/total_tickets)*100, (fora_prazo/total_tickets)*100, (sem_prazo/total_tickets)*100]
        }).to_excel(writer, sheet_name='Resumo Prazos', index=False)
        ranking_clientes.to_excel(writer, sheet_name='Ranking Clientes', index=False)
        
        # Adiciona a aba de resumo da qualidade, se a análise foi feita
        if "resumo_df" in qualidade_results:
            qualidade_results["resumo_df"].to_excel(writer, sheet_name='Resumo Qualidade', index=False)

    return (df, output_file, graph_data_prazo, no_prazo, fora_prazo, sem_prazo, 
            ranking_clientes, qualidade_results)


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

        try:
            (df, output_file, graph_data_prazo, no_prazo, fora_prazo, sem_prazo, 
             ranking, qualidade_results) = process_excel(file_path)

            return render_template('result.html', 
                                   graph_data_prazo=graph_data_prazo,
                                   no_prazo=no_prazo, fora_prazo=fora_prazo, sem_prazo=sem_prazo,
                                   ranking=ranking.to_html(classes="table table-bordered table-striped", index=False),
                                   qualidade_results=qualidade_results,
                                   df_preview=df.head(10).to_html(classes="table table-bordered table-striped", index=False)
                                  )
        except Exception as e:
            return f"Ocorreu um erro ao processar o arquivo: <br>{e}"

    return render_template('upload.html')


@app.route('/download')
def download_file():
    return send_file(os.path.join(UPLOAD_FOLDER, 'output.xlsx'), as_attachment=True)


# As outras rotas como /fora_do_prazo não foram alteradas e podem permanecer as mesmas
# ... (cole a sua rota /fora_do_prazo aqui se desejar)


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(host='0.0.0.0', port=5000, debug=True)