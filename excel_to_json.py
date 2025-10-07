#!/usr/bin/env python3
"""
Script para converter dados do Excel para JSON para uso no dashboard
"""

import pandas as pd
import numpy as np
import json
from datetime import datetime

def converter_excel_para_json(excel_path, output_path):
    """
    Converte dados do Excel para JSON formatado para o dashboard
    
    Args:
        excel_path (str): Caminho para o arquivo Excel
        output_path (str): Caminho para salvar o arquivo JSON
    """
    print(f"Convertendo {excel_path} para JSON...")
    
    # Carregar os dados do Excel
    compras_df = pd.read_excel(excel_path, sheet_name='COMPRAS')
    mb51_df = pd.read_excel(excel_path, sheet_name='MB51 ')
    
    # Calcular métricas principais
    total_notas_mb51 = len(mb51_df)
    total_notas_ajustadas = len(compras_df)
    percentual_ajustadas = (total_notas_ajustadas / total_notas_mb51) * 100
    
    # Calcular tempo médio de ajuste
    compras_df['Data Solicitação de Ajuste'] = pd.to_datetime(compras_df['Data Solicitação de Ajuste'], errors='coerce')
    compras_df['Data Ajuste'] = pd.to_datetime(compras_df['Data Ajuste'], errors='coerce')
    compras_df['Tempo de Ajuste'] = (compras_df['Data Ajuste'] - compras_df['Data Solicitação de Ajuste']).dt.days
    tempo_medio_ajuste = compras_df['Tempo de Ajuste'].mean()
    tempo_mediano_ajuste = compras_df['Tempo de Ajuste'].median()
    tempo_min_ajuste = compras_df['Tempo de Ajuste'].min()
    tempo_max_ajuste = compras_df['Tempo de Ajuste'].max()
    
    # Obter distribuição por categoria
    categorias_labels = compras_df['Categoria'].value_counts().index.tolist()
    categorias_values = compras_df['Categoria'].value_counts().values.tolist()
    
    # Obter distribuição por tipo de ajuste
    if 'Tipo Ajuste Solicitado' in compras_df.columns:
        tipos_ajuste_labels = compras_df['Tipo Ajuste Solicitado'].value_counts().head(8).index.tolist()
        tipos_ajuste_values = compras_df['Tipo Ajuste Solicitado'].value_counts().head(8).values.tolist()
    else:
        tipos_ajuste_labels = []
        tipos_ajuste_values = []
    
    # Obter top fornecedores
    fornecedores_labels = compras_df['Fornecedor'].value_counts().head(5).index.tolist()
    fornecedores_values = compras_df['Fornecedor'].value_counts().head(5).values.tolist()
    
    # Obter top compradores
    compradores_labels = compras_df['Comprador'].value_counts().head(5).index.tolist()
    compradores_values = compras_df['Comprador'].value_counts().head(5).values.tolist()
    
    # Extrair informações de pendências do PDF (valores mencionados)
    pendencias_historico = {
        'labels': ['Junho', 'Julho', 'Setembro'],
        'values': [117, 44, 33]
    }
    
    # Extrair informações de envios manuais vs ME do PDF
    envios_manuais = {
        'atual': 37,
        'anterior': 27,
        'variacao_percentual': ((37 - 27) / 27) * 100
    }
    
    # Determinar o período dos dados
    data_min = compras_df['Data Solicitação de Ajuste'].min()
    data_max = compras_df['Data Solicitação de Ajuste'].max()
    periodo = f"{data_min.strftime('%d/%m/%Y') if not pd.isna(data_min) else 'N/A'} a {data_max.strftime('%d/%m/%Y') if not pd.isna(data_max) else 'N/A'}"
    
    # Criar estrutura de dados para o dashboard
    dados = {
        'representatividade': {
            'labels': ['Notas Ajustadas', 'Notas sem Ajuste'],
            'values': [total_notas_ajustadas, total_notas_mb51 - total_notas_ajustadas],
            'colors': ['#003057', '#58595B']
        },
        'categorias': {
            'labels': categorias_labels,
            'values': categorias_values,
            'colors': ['#003057', '#0091DA', '#FF671F', '#009A44', '#58595B']
        },
        'tiposAjuste': {
            'labels': tipos_ajuste_labels,
            'values': tipos_ajuste_values
        },
        'pendencias': pendencias_historico,
        'fornecedores': {
            'labels': fornecedores_labels,
            'values': fornecedores_values
        },
        'compradores': {
            'labels': compradores_labels,
            'values': compradores_values
        },
        'tempoAjuste': {
            'min': float(tempo_min_ajuste) if not pd.isna(tempo_min_ajuste) else 0,
            'max': float(tempo_max_ajuste) if not pd.isna(tempo_max_ajuste) else 0,
            'media': float(tempo_medio_ajuste) if not pd.isna(tempo_medio_ajuste) else 0,
            'mediana': float(tempo_mediano_ajuste) if not pd.isna(tempo_mediano_ajuste) else 0
        },
        'metricas': {
            'total_notas_mb51': int(total_notas_mb51),
            'total_notas_ajustadas': int(total_notas_ajustadas),
            'percentual_ajustadas': round(float(percentual_ajustadas), 1),
            'tempo_medio_ajuste': round(float(tempo_medio_ajuste), 1) if not pd.isna(tempo_medio_ajuste) else 0,
            'tempo_mediano_ajuste': round(float(tempo_mediano_ajuste), 1) if not pd.isna(tempo_mediano_ajuste) else 0,
            'pendencias_atuais': 33
        },
        'envios_manuais': envios_manuais,
        'periodo': periodo
    }
    
    # Salvar como JSON
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)
    
    print(f"Dados convertidos e salvos em {output_path}")
    print(f"Total de notas no período (MB51): {total_notas_mb51}")
    print(f"Total de notas que passaram por ajuste: {total_notas_ajustadas}")
    print(f"Percentual de representatividade: {percentual_ajustadas:.1f}%")
    print(f"Tempo médio de ajuste: {tempo_medio_ajuste:.1f} dias")
    print(f"Período dos dados: {periodo}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 2:
        excel_path = sys.argv[1]
        output_path = sys.argv[2]
    else:
        excel_path = "../upload/BDCompilado.xlsx"
        output_path = "sample_data.json"
    
    converter_excel_para_json(excel_path, output_path)
