from flask import Flask, render_template, request, send_file, flash
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook


app = Flask(__name__)
app.secret_key = 'supersecretkey'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    global df
    file = request.files['file']
    df = pd.read_excel(file)

    df["Mês"] = pd.to_datetime(df["Mês"], format='%b/%y', errors='coerce').dt.strftime('%b/%y')

    df["Audiências Realizadas"] = df["Audiências Realizadas"] + df["Audiências realizadas no CEJUSC"]
    df["Despachos"] = df["Despachos"] + df["Despachos no CEJUSC"]
    df["Decisões"] = df["Decisões"] + df["Decisões no CEJUSC"]
    df["Julgamentos com Mérito"] = df["Julgamentos com Mérito"] + df["Julgamentos com mérito no CEJUSC"]
    df["Julgamentos sem Mérito"] = df["Julgamentos sem Mérito"] + df["Julgamentos sem mérito no CEJUSC"]

    df = df.drop(["Audiências realizadas no CEJUSC", "Despachos no CEJUSC", "Decisões no CEJUSC", "Julgamentos com mérito no CEJUSC", "Julgamentos sem mérito no CEJUSC"], axis=1)

    global soma_por_mes_ano
    soma_por_mes_ano = df.groupby('Mês')['Excesso de Prazo Sentença'].sum().reset_index()
    soma_por_mes_ano.rename(columns={'Excesso de Prazo Sentença': 'Soma Excesso de Prazo'}, inplace=True)
    soma_por_mes_ano.insert(0, 'Mês_Duplicado', soma_por_mes_ano['Mês'])
    soma_por_mes_ano['Mês_Duplicado'] = pd.to_datetime(soma_por_mes_ano['Mês_Duplicado'], format='%b/%y').dt.to_period('M')
    soma_por_mes_ano = soma_por_mes_ano.sort_values(by='Mês_Duplicado')
    soma_por_mes_ano = soma_por_mes_ano.drop_duplicates(subset=['Mês'], keep='first')
    soma_por_mes_ano.insert(0, 'Magistrado', df['Magistrado'])
    soma_por_mes_ano = soma_por_mes_ano[['Magistrado', 'Mês', 'Soma Excesso de Prazo']]

    soma_html = soma_por_mes_ano.to_html(classes='table table-striped', index=False)
    return render_template('show_table.html', tables=[soma_html])

@app.route('/download', methods=['POST'])
def download():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Planilha1', index=False)
        soma_por_mes_ano.to_excel(writer, sheet_name='Planilha2', index=False)

    output.seek(0)

    return send_file(output, as_attachment=True, download_name='output.xlsx')

if __name__ == '__main__':
    app.run(debug=True)



