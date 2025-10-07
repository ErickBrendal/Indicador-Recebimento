# Dashboard Interativo de Recebimento Fiscal

Este dashboard interativo foi desenvolvido para visualizar e analisar os dados de recebimento fiscal da Saint-Gobain, com foco em indicadores de desempenho, ajustes e tendências.

## Instruções para Execução

### Execução Local

Para executar o dashboard localmente, siga estas etapas simples:

1. Abra o arquivo `index.html` em qualquer navegador moderno (Chrome, Firefox, Edge, Safari)
2. Não é necessário servidor web ou dependências adicionais para visualização básica

### Funcionalidades

O dashboard inclui as seguintes funcionalidades:

- **Navegação por seções**: Acesse diferentes visualizações através do menu lateral
- **Filtros interativos**: Filtre os dados por período, fornecedor, categoria e depósito
- **Gráficos interativos**: Passe o mouse sobre os elementos dos gráficos para ver detalhes
- **Exportação para PDF**: Utilize o botão "Exportar PDF" para gerar uma versão para impressão
- **Exportação para PPT**: Funcionalidade simulada (seria implementada com PptxGenJS)

## Estrutura do Projeto

- `index.html`: Arquivo principal contendo todo o código HTML, CSS e JavaScript
- `README.md`: Este arquivo de documentação

## Conversão de Dados

Os dados foram extraídos do arquivo Excel `BDCompilado.xlsx` e convertidos para o formato JSON para uso no dashboard. O processo de conversão pode ser realizado com o seguinte script Python:

```python
import pandas as pd
import json

# Carregar os dados do Excel
excel_path = 'BDCompilado.xlsx'
compras_df = pd.read_excel(excel_path, sheet_name='COMPRAS')
mb51_df = pd.read_excel(excel_path, sheet_name='MB51 ')

# Criar estrutura de dados para o dashboard
dados = {
    'representatividade': {
        'labels': ['Notas Ajustadas', 'Notas sem Ajuste'],
        'values': [len(compras_df), len(mb51_df) - len(compras_df)]
    },
    'categorias': {
        'labels': compras_df['Categoria'].value_counts().index.tolist(),
        'values': compras_df['Categoria'].value_counts().values.tolist()
    },
    # Adicionar outros dados conforme necessário
}

# Salvar como JSON
with open('dados_dashboard.json', 'w') as f:
    json.dump(dados, f, indent=2)
```

## Personalização

Para personalizar o dashboard:

1. **Cores**: Modifique as variáveis CSS no início do arquivo `index.html`
2. **Dados**: Atualize o objeto `dados` no script JavaScript
3. **Filtros**: Implemente a lógica de filtragem real no evento `change` dos elementos `form-select`

## Exportação para PowerPoint

Para implementar a exportação para PowerPoint, seria necessário:

1. Utilizar a biblioteca PptxGenJS (já incluída nos scripts)
2. Capturar os gráficos como imagens usando html2canvas
3. Criar slides com as imagens capturadas e os dados relevantes

Exemplo de implementação:

```javascript
document.getElementById('exportPpt').addEventListener('click', function() {
  // Criar nova apresentação
  let pptx = new PptxGenJS();
  
  // Adicionar slide de capa
  let slide = pptx.addSlide();
  slide.addText("Recebimento Fiscal — Resultados e Produtividade", { 
    x: 1, y: 1, w: 8, h: 1, fontSize: 24, color: "003057" 
  });
  
  // Capturar e adicionar gráficos
  html2canvas(document.getElementById('representatividadeChart')).then(canvas => {
    slide = pptx.addSlide();
    slide.addImage({ data: canvas.toDataURL(), x: 1, y: 1, w: 8, h: 4 });
  });
  
  // Salvar arquivo
  pptx.writeFile("Recebimento_Fiscal.pptx");
});
```

## Compatibilidade

O dashboard é compatível com:

- Chrome 88+
- Firefox 85+
- Edge 88+
- Safari 14+

## Contato

Para suporte ou dúvidas, entre em contato com:
- Email: analista.fiscal@empresa.com
- Ramal: 1234
