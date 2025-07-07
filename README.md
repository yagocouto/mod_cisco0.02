# Processamento de Dados de Interfaces Cisco

Este projeto consiste em um script Python para processar arquivos de texto gerados por equipamentos Cisco, extrair informações relevantes sobre interfaces de rede e dispositivos vizinhos via CDP (Cisco Discovery Protocol), e gerar relatórios em formato Excel.

---

## Funcionalidades

- Leitura de arquivos `.txt` contendo saídas de comandos Cisco.
- Extração do nome do dispositivo (hostname).
- Extração do status das interfaces, incluindo:
  - Interface local, descrição, status, VLAN, duplex, velocidade e tipo.
- Extração de informações CDP dos dispositivos vizinhos, como:
  - Nome do dispositivo vizinho, IP e porta.
- Detecção de erros relevantes nas interfaces (input errors e CRC errors).
- Geração de planilhas Excel com as informações extraídas, criando ou atualizando abas conforme o arquivo processado.

---

## Estrutura do Código

- `ler_arquivo(caminho)`: Lê o arquivo texto e retorna suas linhas.
- `extrair_valor_unico(chave, linha)`: Extrai um valor único baseado em uma chave de texto.
- `extrair_ip_vizinho_cdp(linhas, indice_inicial)`: Obtém o IP do vizinho CDP em um bloco de linhas.
- `extrair_entradas_cdp(linhas)`: Extrai os dados dos dispositivos vizinhos via CDP.
- `extrair_bloco_interface(linhas, interface_desejada)`: Captura o bloco de texto referente a uma interface específica.
- `extrair_observacao_erros(linhas, interface_desejada)`: Busca por erros de entrada nas interfaces.
- `extrair_interfaces_status(linhas, device_name)`: Monta a lista de interfaces com seus respectivos dados.
- `gerar_excel(dados, caminho_excel, nome_aba)`: Gera ou atualiza o arquivo Excel com os dados extraídos.
- `processar_mod_cisco()`: Função principal que realiza a leitura dos arquivos na pasta `entrada`, processa os dados e salva o resultado na pasta `saida`.

---

## Requisitos

- Python 3.7 ou superior
- Bibliotecas Python:
  - `pandas`
  - `openpyxl`

Para instalar as dependências:

```bash
pip install pandas openpyxl
