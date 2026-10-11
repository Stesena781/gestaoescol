import pandas as pd

ABAS = {
    "classificacao": "Classificação",
    "escolhas": "Escolha dos Alunos",
    "vagas": "Vagas",
}


def normalizar_opm(valor):
    """Unifica '5° BAEP' / '5º BAEP' e remove espaços duplicados."""
    if pd.isna(valor):
        return None
    s = str(valor).strip().replace("°", "º")
    s = " ".join(s.split())
    return s.upper()


def carregar_classificacao(caminho):
    df = pd.read_excel(caminho, sheet_name=ABAS["classificacao"])
    df.columns = [str(c).strip() for c in df.columns]
    df["CLASS_NUM"] = df["CLASS"].astype(str).str.extract(r"(\d+)").astype(int)
    df["RE/RG"] = df["RE/RG"].astype(str).str.strip().str.upper()
    df["Remanescente"] = df["Remanescente"].astype(str).str.strip().str.upper()
    df["NOME"] = df["NOME"].astype(str).str.strip()
    df["PRIORIDADE"] = df["Remanescente"].eq("SIM").map({True: 0, False: 1})
    return df.sort_values(["PRIORIDADE", "CLASS_NUM"]).reset_index(drop=True)


def carregar_escolhas(caminho, n_opcoes=58):
    df = pd.read_excel(caminho, sheet_name=ABAS["escolhas"])
    df.columns = [str(c).strip() for c in df.columns]
    df["RE/RG"] = df["RE/RG"].astype(str).str.strip().str.upper()
    for i in range(1, n_opcoes + 1):
        col = f"OPM{i}"
        if col in df.columns:
            df[col] = df[col].apply(normalizar_opm)
    return df


def carregar_vagas(caminho):
    df = pd.read_excel(caminho, sheet_name=ABAS["vagas"])
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={df.columns[0]: "OPM", df.columns[1]: "VAGAS"})
    df["OPM"] = df["OPM"].apply(normalizar_opm)
    df["VAGAS"] = pd.to_numeric(df["VAGAS"], errors="coerce")
    df = df.dropna(subset=["OPM", "VAGAS"])
    df = df[~df["OPM"].str.startswith(("TOTAL", "ESCOLHA", "VAGAS"))]
    df["VAGAS"] = df["VAGAS"].astype(int)
    return df.reset_index(drop=True)


def alocar(classificacao, escolhas, vagas):
    saldo = dict(zip(vagas["OPM"], vagas["VAGAS"]))
    escolhas_idx = escolhas.set_index("RE/RG")
    colunas_opm = [c for c in escolhas.columns if c.startswith("OPM")]

    registros = []
    for _, aluno in classificacao.iterrows():
        rg = aluno["RE/RG"]
        opcao = None

        if rg in escolhas_idx.index:
            linha = escolhas_idx.loc[rg]
            if isinstance(linha, pd.DataFrame):
                linha = linha.iloc[0]
            for col in colunas_opm:
                opm = linha[col]
                if opm is None or opm not in saldo:
                    continue
                if saldo[opm] > 0:
                    opcao = opm
                    saldo[opm] -= 1
                    break

        registros.append({
            "Classificação": aluno["CLASS_NUM"],
            "RG": rg,
            "Nome": aluno["NOME"],
            "Remanescente": aluno["Remanescente"],
            "Opção Alocada": opcao if opcao else "SEM ALOCAÇÃO",
            "Ordem de Processamento": len(registros) + 1,
        })

    resultado = pd.DataFrame(registros)
    sobra = pd.DataFrame(
        [{"OPM": k, "Vagas Restantes": v} for k, v in saldo.items()]
    ).sort_values("OPM").reset_index(drop=True)
    return resultado, sobra