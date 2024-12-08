import os
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'secret_key'  # Para mensagens flash
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'

# Criar diretórios para uploads e arquivos processados
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

COST_PER_POINT = 0.403  # Exemplo para custo por ponto (não usado diretamente aqui)


def process_excel(file_path):
    """Processar o arquivo Excel enviado pelo usuário."""
    # Carregar os dados da aba "Livelo"
    df = pd.read_excel(file_path, sheet_name='livelo')

    # Filtrar pedidos com o cupom "Livelo" (case insensitive)
    if 'Coupon' not in df.columns:
        raise ValueError("A coluna 'Coupon' não foi encontrada na planilha.")
    df = df[df['Coupon'].str.contains('livelo', case=False, na=False)]

    # Verificar se as colunas necessárias estão presentes
    required_columns = ['Order', 'SKU Selling Price', 'Quantity_SKU', 'CPP (custo por ponto']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"As seguintes colunas estão faltando na tabela: {', '.join(missing_columns)}")

    # Converter as colunas numéricas para float (caso necessário)
    df['SKU Selling Price'] = pd.to_numeric(df['SKU Selling Price'], errors='coerce')
    df['Quantity_SKU'] = pd.to_numeric(df['Quantity_SKU'], errors='coerce')
    df['CPP (custo por ponto'] = pd.to_numeric(df['CPP (custo por ponto'], errors='coerce')

    # Calcular o valor total do SKU Selling Price com base na quantidade
    df['Total SKU Selling Price'] = df['SKU Selling Price'] * df['Quantity_SKU']

    # Agrupar por número do pedido (Order) e calcular o subtotal
    grouped = df.groupby('Order', as_index=False).agg({
        'Total SKU Selling Price': 'sum',  # Subtotal por pedido
        'CPP (custo por ponto': 'first'  # Custo por ponto (assumindo que é constante por pedido)
    })

    # Renomear colunas para facilitar a leitura
    grouped.rename(columns={'Total SKU Selling Price': 'Subtotal'}, inplace=True)

    # Calcular os pontos Livelo (3 pontos por R$1 no subtotal)
    grouped['Pontos Livelo'] = grouped['Subtotal'] * 3

    # Calcular o custo total dos pontos
    grouped['Custo Total'] = grouped['Pontos Livelo'] * grouped['CPP (custo por ponto']

    return grouped



@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verificar se o arquivo foi enviado
        if 'file' not in request.files:
            flash('Nenhum arquivo foi enviado!', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            flash('Nenhum arquivo selecionado!', 'error')
            return redirect(request.url)
        
        # Salvar o arquivo enviado
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        try:
            # Processar o arquivo
            processed_data = process_excel(file_path)
            
            # Salvar o arquivo processado
            output_file = os.path.join(app.config['PROCESSED_FOLDER'], secure_filename(f'processed_{file.filename}'))
            processed_data.to_excel(output_file, index=False)
            
            # Exibir resultado e permitir download
            return render_template('index.html', 
                                   table=processed_data.to_html(classes='table table-striped', index=False), 
                                   download_link=os.path.basename(output_file))
        except Exception as e:
            flash(f'Erro ao processar o arquivo: {str(e)}', 'error')
            return redirect(request.url)
    
    return render_template('index.html')


@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    # Construir o caminho absoluto do arquivo processado
    filepath = os.path.join(app.config['PROCESSED_FOLDER'], filename)
    print(f"Tentando localizar o arquivo: {filepath}")  # Log para depuração
    if not os.path.exists(filepath):
        flash("Arquivo não encontrado para download.", "error")
        return redirect(url_for('index'))
    return send_file(filepath, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
