import os
import pandas as pd
from docx import Document

def extrair_info(documento):
    dados = []
    tabela = documento.tables[0]

    for i, linha in enumerate(tabela.rows[1:], start=1):
        celulas = linha.cells
        data = celulas[0].text.strip()
        funcionario = celulas[1].text.strip()
        horas = celulas[2].text.strip()
        percentual = celulas[3].text.strip()

        if percentual.lower() == 'folga':
            perc = 'Folga'
            horas_trabalhadas = 0
        else:
            perc = percentual.replace('%', '').strip()
            horas_trabalhadas = float(horas.replace(',', '.'))

        dados.append({
            'Data': data,
            'Funcionário': funcionario,
            'Horas': horas_trabalhadas,
            'Percentual': perc
        })

    return dados

def main():
    caminho_docx = os.path.join('dados', 'horas_ms_0307.docx')
    doc = Document(caminho_docx)
    registros = extrair_info(doc)

    df = pd.DataFrame(registros)

    resumo = df.pivot_table(
        index='Funcionário',
        columns='Percentual',
        values='Horas',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    colunas_ordenadas = ['Funcionário', '50', '100', 'Folga']
    for col in colunas_ordenadas:
        if col not in resumo.columns:
            resumo[col] = 0
    resumo = resumo[colunas_ordenadas]

    os.makedirs('relatorios', exist_ok=True)
    resumo.to_csv(os.path.join('relatorios', 'resumo_horas.csv'), index=False, sep=';')
    print("✅ Relatório gerado com sucesso em 'relatorios/resumo_horas.csv'.")

if __name__ == "__main__":
    main()
