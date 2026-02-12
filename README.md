# Automação de Briefing 

Este projeto automatiza a geração de planilhas de planejamento, campanha e expert a partir de um arquivo de "Briefing" base.

## Estrutura do Projeto

*   `Planilha Base/`: Coloque aqui os arquivos de briefing (`.xlsx`) que deseja processar. O script processará todos os arquivos encontrados nesta pasta.
*   `Output/`: As planilhas geradas serão salvas automaticamente nesta pasta.
*   `src/`: Contém o código fonte (script Python).
*   `requirements.txt`: Lista de dependências do Python necessárias.

## Pré-requisitos

1.  **Python 3.x** instalado.
2.  Bibliotecas Python listadas em `requirements.txt`.

## Instalação

1.  Abra o terminal (Prompt de Comando ou PowerShell) na pasta do projeto.
2.  Instale as dependências executando o comando:

```bash
pip install -r requirements.txt
```

## Como Usar

1.  **Prepare o Arquivo Base:**
    *   Certifique-se de que o arquivo de briefing (ex: `[BLK1025l] - [FMA] - [BRIEFING] .xlsx`) esteja salvo na pasta `Planilha Base`.
    *   Você pode colocar múltiplos arquivos nesta pasta; o script processará todos eles.

2.  **Execute a Automação:**
    *   No terminal, execute o seguinte comando:

```bash
python src/generate_files.py
```

3.  **Verifique os Resultados:**
    *   Vá até a pasta `Output`.
    *   Você encontrará 3 arquivos gerados para cada briefing processado, seguindo o padrão de nomenclatura:
        *   `[00.00][ID] Planejamento.xlsx`
        *   `[01.00][ID] Campanha 1.xlsx`
        *   `[01.01][ID][Campanha 1] Expert 1.xlsx`

## Funcionalidades

O script lê automaticamente do arquivo base:
*   **Identificação:** ID da Campanha, Nome do Expert.
*   **Datas:** Períodos das fases (Captação, Aquecimento, Lembrete, Remarketing).
*   **Metas:** Custo por Lead (CPL), Metas de Venda e Faturamento.
*   **Investimentos:** Orçamento previsto por plataforma (Meta Ads, Google Ads, YouTube) para cada fase.
*   **Detalhamento Diário:** Dados dia-a-dia da fase de Captação (Investimento, Leads, etc.).

E preenche as planilhas de destino mantendo a formatação e estrutura exigidas.
