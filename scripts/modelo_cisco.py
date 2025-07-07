# Reexecutando após reset: código corrigido e formatado adequadamente

from pathlib import Path
import re
import pandas as pd


def ler_arquivo(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        return f.readlines()


def extrair_valor_unico(chave, linha):
    if chave in linha:
        return linha.split(chave)[1].strip()
    return None


def extrair_ip_vizinho_cdp(linhas, indice_inicial):
    for i in range(indice_inicial, len(linhas)):
        linha = linhas[i].strip()
        if linha.startswith("IP address:"):
            return linha.split("IP address:")[1].strip()
        if linha and not linha.startswith(" "):
            break
    return ""


def extrair_entradas_cdp(linhas):
    entradas = []
    for i, linha in enumerate(linhas):
        if "Device ID:" in linha:
            neighbor_device = extrair_valor_unico("Device ID:", linha)
            neighbor_port = ""
            local_port = ""
            ip = ""

            for j in range(i + 1, len(linhas)):
                if "Interface:" in linhas[j]:
                    partes = linhas[j].split(",")
                    if len(partes) >= 2:
                        local_port = partes[0].split("Interface:")[1].strip()
                        neighbor_port = (
                            partes[1].split("Port ID (outgoing port):")[1].strip()
                        )
                elif "IP address:" in linhas[j]:
                    ip = linhas[j].split("IP address:")[1].strip()
                elif linhas[j].strip() == "":
                    break

            entradas.append(
                {
                    "Local Port": local_port.replace("GigabitEthernet", "Gi").replace(
                        "FastEthernet", "Fa"
                    ),
                    "Neighbor Device": neighbor_device,
                    "Neighbor IP": ip,
                    "Neighbor Port": neighbor_port,
                }
            )
    return entradas


def extrair_bloco_interface(linhas, interface_desejada):
    bloco = []
    capturando = False

    for linha in linhas:
        if linha.startswith("GigabitEthernet") or linha.startswith("FastEthernet"):
            if capturando:
                break
            if linha.startswith(interface_desejada):
                capturando = True
                bloco.append(linha)
        elif capturando:
            bloco.append(linha)

    return bloco


def extrair_observacao_erros(linhas, interface_desejada):
    bloco = extrair_bloco_interface(linhas, interface_desejada)
    for linha in bloco:
        if "input errors" in linha and "CRC" in linha:
            match = re.search(r"(\d+)\s+input errors, (\d+)\s+CRC", linha)
            if match:
                input_errors = int(match.group(1))
                crc_errors = int(match.group(2))
                if input_errors >= 5 or crc_errors >= 5:
                    return linha.strip()
    return ""


def extrair_interfaces_status(linhas, device_name):
    interfaces = []
    capturar = False

    for linha in linhas:
        if linha.strip().startswith("Port") and linha.strip().endswith("Type"):
            capturar = True
            continue

        if capturar:
            if re.match(r"^\S+#$", linha.strip()):
                continue
            if "show interfaces status err-disabled" in linha.lower():
                break
            port = linha[0:10].strip()
            name = linha[10:29].strip()
            status = linha[29:42].strip()
            vlan = linha[42:53].strip()
            duplex = linha[53:60].strip()
            speed = linha[60:67].strip()
            tipo = linha[67:].strip()

            interfaces.append(
                {
                    "Device Name": device_name,
                    "Interface": port,
                    "Description": name,
                    "CDP Device ID": "",
                    "LLDP Device ID": "",
                    "Neighbor Dest. Port": "",
                    "CDP Neighbor IP Address": "",
                    "Status": status,
                    "Vlan": vlan,
                    "Duplex": duplex,
                    "Speed": speed,
                    "Type": tipo,
                    "Observação": "",
                    "Destination Port": "",
                }
            )
    return interfaces


def gerar_excel(dados, caminho_excel, nome_aba):
    df = pd.DataFrame(dados)
    caminho_excel = Path(caminho_excel)

    if caminho_excel.exists():
        with pd.ExcelWriter(
            caminho_excel, engine="openpyxl", mode="a", if_sheet_exists="overlay"
        ) as writer:
            workbook = writer.book
            if nome_aba in workbook.sheetnames:
                del workbook[nome_aba]
            df.to_excel(writer, sheet_name=nome_aba, index=False)
    else:
        with pd.ExcelWriter(caminho_excel, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)

    print(f"Planilha salva: {caminho_excel} | Aba: {nome_aba}")


def processar_mod_cisco():
    pasta = Path("entrada")
    arquivos_txt = list(pasta.glob("*.txt"))
    caminho_excel = "saida/interfaces_cisco.xlsx"

    for arquivo in arquivos_txt:
        print(f"Lendo arquivo: {arquivo.name}")
        nome_aba = arquivo.stem
        caminho_arquivo = arquivo.resolve()

        linhas = ler_arquivo(caminho_arquivo)
        device_name = ""
        for linha in linhas:
            if linha.startswith("hostname"):
                device_name = linha.split("hostname")[1].strip()
                break

        interfaces = extrair_interfaces_status(linhas, device_name)
        cdp_entries = extrair_entradas_cdp(linhas)

        for interface in interfaces:
            for entry in cdp_entries:
                if interface["Interface"] == entry["Local Port"]:
                    interface["CDP Device ID"] = entry["Neighbor Device"]
                    interface["CDP Neighbor IP Address"] = entry["Neighbor IP"]
                    interface["Neighbor Dest. Port"] = entry["Neighbor Port"]

            nome_interface = (
                interface["Interface"]
                .replace("Gi", "GigabitEthernet")
                .replace("Fa", "FastEthernet")
            )
            observacao = extrair_observacao_erros(linhas, nome_interface)
            if observacao:
                interface["Observação"] = observacao

        gerar_excel(interfaces, caminho_excel, nome_aba[:30])
