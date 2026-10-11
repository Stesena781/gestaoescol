import os
import pandas as pd
from flask import Flask, render_template, request, send_file

from alocador import (
    carregar_classificacao,
    carregar_escolhas,
    carregar_vagas,
    alocar,
)

app = Flask(__name__)

UPLOADS = "uploads"
os.makedirs(UPLOADS, exist_ok=True)


def salvar(campo):
    arquivo = request.files.get(campo)

    if not arquivo or arquivo.filename == "":
        raise ValueError(
            f"Arquivo ausente no campo '{campo}'"
        )

    destino = os.path.join(
        UPLOADS,
        campo + ".xlsx"
    )

    arquivo.save(destino)

    return destino


def gerar_prova_real(resultado, classificacao):

    base = resultado.merge(
        classificacao[
            [
                "RE/RG",
                "SEXO"
            ]
        ],
        left_on="RG",
        right_on="RE/RG",
        how="left"
    )

    base = base[
        base["Opção Alocada"] != "SEM ALOCAÇÃO"
    ]

    prova = (
        base
        .groupby(
            ["Opção Alocada", "SEXO"]
        )
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    prova = prova.rename(
        columns={
            "Opção Alocada": "OPM",
            "M": "MAS",
            "F": "FEM"
        }
    )

    if "MAS" not in prova.columns:
        prova["MAS"] = 0

    if "FEM" not in prova.columns:
        prova["FEM"] = 0

    prova = prova.sort_values(
        "OPM"
    ).reset_index(drop=True)

    return prova


def gerar_excel(
    resultado,
    sobra,
    vagas,
    classificacao,
    caminho_saida
):

    ultimo = (
        resultado[
            resultado["Opção Alocada"]
            != "SEM ALOCAÇÃO"
        ]
        .groupby("Opção Alocada")[
            "Classificação"
        ]
        .max()
        .reset_index()
        .rename(
            columns={
                "Opção Alocada": "OPM",
                "Classificação": "Ultimo Classificado"
            }
        )
    )

    mapa_sobra = dict(
        zip(
            sobra["OPM"],
            sobra["Vagas Restantes"]
        )
    )

    resumo = vagas.merge(
        ultimo,
        on="OPM",
        how="left"
    )

    resumo["Vagas Restantes"] = (
        resumo["OPM"]
        .map(mapa_sobra)
        .fillna(0)
        .astype(int)
    )

    resumo["Alocadas"] = (
        resumo["VAGAS"]
        - resumo["Vagas Restantes"]
    )

    prova_real = gerar_prova_real(
        resultado,
        classificacao
    )

    with pd.ExcelWriter(
        caminho_saida,
        engine="openpyxl"
    ) as writer:

        resultado.to_excel(
            writer,
            sheet_name="Alocação",
            index=False
        )

        resumo.to_excel(
            writer,
            sheet_name="Resumo por OPM",
            index=False
        )

        sobra.to_excel(
            writer,
            sheet_name="Vagas Restantes",
            index=False
        )

        prova_real.to_excel(
            writer,
            sheet_name="PROVA REAL",
            index=False
        )

    return caminho_saida


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        try:

            cls = carregar_classificacao(
                salvar("classificacao")
            )

            esc = carregar_escolhas(
                salvar("escolhas")
            )

            vag = carregar_vagas(
                salvar("vagas")
            )

            resultado, sobra = alocar(
                cls,
                esc,
                vag
            )

            saida = os.path.join(
                UPLOADS,
                "resultado.xlsx"
            )

            gerar_excel(
                resultado,
                sobra,
                vag,
                cls,
                saida
            )

            return send_file(
                saida,
                as_attachment=True,
                download_name="alocacao_final.xlsx"
            )

        except Exception as erro:

            return render_template(
                "index.html",
                erro=str(erro)
            ), 400

    return render_template(
        "index.html"
    )


if __name__ == "__main__":
    app.run(
        debug=True
    )
