# Robô lançador de GNRE no SPED Fiscal

Lê os PDFs das guias (GNRE, DARE-SP, DAE-MG e Guia de Recolhimento do Paraná), encontra o CT-e de cada uma no TXT do SPED Fiscal
(EFD ICMS/IPI) e gera um TXT novo com os registros **D195** (GNRE – GUIA DE RECOLHIMENTO) e
**D197** (ajuste **MT10001002**, "NUMERO DA GUIA ..." e o valor principal, sem juros/multa).
DAR-1 de MT é ignorado e só aparece no relatório.

## Instalação (uma vez)

1. Ter o Python instalado (com "Add python.exe to PATH").
2. Criar a pasta `LANCAR GNRE` na Área de Trabalho.
3. Baixar o arquivo [`LANCAR GNRE.bat`](https://github.com/NetoPianta/robo-gnre/raw/main/LANCAR%20GNRE.bat) para dentro dessa pasta.
4. Dar dois cliques no `.bat`. Na primeira vez ele baixa o robô e instala os componentes (1 a 2 minutos).

## Uso (todo mês)

1. Dois cliques em `LANCAR GNRE.bat` — ele sempre pega a versão mais nova daqui antes de rodar.
2. Escolher o TXT original do SPED do mês e depois a pasta com os PDFs (`GNRE <nº do CT-e>.pdf`).
3. Importar no PVA o arquivo que termina em ` - COM GNRE.txt` e conferir o relatório `.csv`.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `LANCAR GNRE.bat` | Atualiza o robô, instala os componentes e roda |
| `gnre_sped.py` | O robô |
| `requirements.txt` | Componentes do Python que o robô usa |

Nenhum dado da empresa (SPED, guias, relatórios) fica neste repositório — só o código.
