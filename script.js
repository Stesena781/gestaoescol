// ======================================================
// SISTEMA DE ESCOLHA
// LEITURA E PROCESSAMENTO DAS PLANILHAS
// ======================================================


// ======================================================
// 1. MOSTRAR NOME DOS ARQUIVOS
// ======================================================

function mostrarArquivo(inputId, nomeId) {

    const input = document.getElementById(inputId);
    const nome = document.getElementById(nomeId);

    input.addEventListener("change", function () {

        if (input.files.length > 0) {

            nome.textContent = input.files[0].name;
            nome.style.color = "#16a34a";

        } else {

            nome.textContent = "Nenhum arquivo selecionado";
            nome.style.color = "#999";

        }

    });
}


mostrarArquivo("classificacao", "classificacao-name");
mostrarArquivo("opcoes", "opcoes-name");
mostrarArquivo("vagas", "vagas-name");



// ======================================================
// 2. LER ARQUIVO EXCEL
// ======================================================

function lerExcel(file) {

    return new Promise((resolve, reject) => {

        const reader = new FileReader();

        reader.onload = function (evento) {

            try {

                const dados = new Uint8Array(
                    evento.target.result
                );

                const workbook = XLSX.read(
                    dados,
                    {
                        type: "array"
                    }
                );

                // Pega a primeira aba
                const primeiraAba =
                    workbook.SheetNames[0];

                const worksheet =
                    workbook.Sheets[primeiraAba];

                // Converte para JSON
                const json =
                    XLSX.utils.sheet_to_json(
                        worksheet,
                        {
                            defval: ""
                        }
                    );

                resolve(json);

            } catch (erro) {

                reject(erro);

            }

        };

        reader.onerror = function () {

            reject(
                new Error("Não foi possível ler o arquivo.")
            );

        };

        reader.readAsArrayBuffer(file);

    });

}


// ======================================================
// 3. NORMALIZAR TEXTO
// ======================================================

function normalizarTexto(valor) {

    if (valor === undefined || valor === null) {
        return "";
    }

    return String(valor)
        .trim()
        .toUpperCase();

}


// ======================================================
// 4. LOCALIZAR COLUNA
// ======================================================

function encontrarColuna(linha, possibilidades) {

    const colunas =
        Object.keys(linha);

    for (const possibilidade of possibilidades) {

        const encontrada =
            colunas.find(coluna =>

                normalizarTexto(coluna)
                    .replace(/\s/g, "")
                    .includes(
                        normalizarTexto(possibilidade)
                            .replace(/\s/g, "")
                    )

            );

        if (encontrada) {

            return encontrada;

        }

    }

    return null;

}


// ======================================================
// 5. PROCESSAR CLASSIFICAÇÃO
// ======================================================

function processarClassificacao(dados) {

    if (!dados || dados.length === 0) {

        throw new Error(
            "A planilha de classificação está vazia."
        );

    }


    const exemplo = dados[0];


    const colunaClassificacao =
        encontrarColuna(
            exemplo,
            [
                "CLASSIFICAÇÃO",
                "CLASSIFICACAO",
                "CLASS"
            ]
        );


    const colunaRG =
        encontrarColuna(
            exemplo,
            [
                "RG/RE",
                "RG",
                "RE"
            ]
        );


    const colunaNome =
        encontrarColuna(
            exemplo,
            [
                "NOME"
            ]
        );


    const colunaSexo =
        encontrarColuna(
            exemplo,
            [
                "SEXO",
                "SEX"
            ]
        );


    if (!colunaClassificacao ||
        !colunaRG ||
        !colunaNome ||
        !colunaSexo) {

        console.log(
            "Colunas encontradas:",
            Object.keys(exemplo)
        );

        throw new Error(
            "Não consegui identificar todas as colunas da classificação."
        );

    }


    const candidatos =
        dados.map(linha => {

            return {

                classificacao:
                    Number(
                        linha[colunaClassificacao]
                    ),

                rg:
                    normalizarTexto(
                        linha[colunaRG]
                    ),

                nome:
                    normalizarTexto(
                        linha[colunaNome]
                    ),

                sexo:
                    normalizarTexto(
                        linha[colunaSexo]
                    )

            };

        });


    // Ordena pela classificação

    candidatos.sort(
        (a, b) =>
            a.classificacao -
            b.classificacao
    );


    return candidatos;

}


// ======================================================
// 6. PROCESSAR OPÇÕES
// ======================================================

function processarOpcoes(dados) {

    if (!dados || dados.length === 0) {

        throw new Error(
            "A planilha de opções está vazia."
        );

    }


    const exemplo = dados[0];


    const colunaRG =
        encontrarColuna(
            exemplo,
            [
                "RG/RE",
                "RG",
                "RE"
            ]
        );


    const colunaNome =
        encontrarColuna(
            exemplo,
            [
                "NOME"
            ]
        );


    const colunaOpcao1 =
        encontrarColuna(
            exemplo,
            [
                "1ª OPÇÃO",
                "1ª OPCAO",
                "1A OPÇÃO",
                "1A OPCAO",
                "OPÇÃO 1",
                "OPCAO 1"
            ]
        );


    const colunaOpcao2 =
        encontrarColuna(
            exemplo,
            [
                "2ª OPÇÃO",
                "2ª OPCAO",
                "2A OPÇÃO",
                "2A OPCAO",
                "OPÇÃO 2",
                "OPCAO 2"
            ]
        );


    const colunaOpcao3 =
        encontrarColuna(
            exemplo,
            [
                "3ª OPÇÃO",
                "3ª OPCAO",
                "3A OPÇÃO",
                "3A OPCAO",
                "OPÇÃO 3",
                "OPCAO 3"
            ]
        );


    if (!colunaRG) {

        throw new Error(
            "Não encontrei a coluna RG/RE na planilha de opções."
        );

    }


    const opcoes = {};


    dados.forEach(linha => {

        const rg =
            normalizarTexto(
                linha[colunaRG]
            );


        if (!rg) {
            return;
        }


        opcoes[rg] = {

            nome:
                colunaNome
                    ? normalizarTexto(
                        linha[colunaNome]
                    )
                    : "",

            opcao1:
                colunaOpcao1
                    ? normalizarTexto(
                        linha[colunaOpcao1]
                    )
                    : "",

            opcao2:
                colunaOpcao2
                    ? normalizarTexto(
                        linha[colunaOpcao2]
                    )
                    : "",

            opcao3:
                colunaOpcao3
                    ? normalizarTexto(
                        linha[colunaOpcao3]
                    )
                    : ""

        };

    });


    return opcoes;

}


// ======================================================
// 7. PROCESSAR VAGAS
// ======================================================

function processarVagas(dados) {

    if (!dados || dados.length === 0) {

        throw new Error(
            "A planilha de vagas está vazia."
        );

    }


    const exemplo = dados[0];


    const colunaOPM =
        encontrarColuna(
            exemplo,
            [
                "OPM"
            ]
        );


    const colunaMasculino =
        encontrarColuna(
            exemplo,
            [
                "MASCULINO",
                "MASCULINO"
            ]
        );


    const colunaFeminino =
        encontrarColuna(
            exemplo,
            [
                "FEMININO"
            ]
        );


    if (!colunaOPM) {

        throw new Error(
            "Não encontrei a coluna OPM na planilha de vagas."
        );

    }


    const vagas = {};


    dados.forEach(linha => {

        const opm =
            normalizarTexto(
                linha[colunaOPM]
            );


        if (!opm) {
            return;
        }


        let masculino = 0;
        let feminino = 0;


        if (colunaMasculino) {

            masculino =
                Number(
                    linha[colunaMasculino]
                ) || 0;

        }


        if (colunaFeminino) {

            feminino =
                Number(
                    linha[colunaFeminino]
                ) || 0;

        }


        vagas[opm] = {

            masculino:
                masculino,

            feminino:
                feminino

        };

    });


    return vagas;

}


// ======================================================
// 8. VERIFICAR VAGA
// ======================================================

function possuiVaga(vagas, opm, sexo) {

    if (!opm) {
        return false;
    }


    const nomeOPM =
        normalizarTexto(opm);


    if (!vagas[nomeOPM]) {
        return false;
    }


    if (sexo === "M") {

        return vagas[nomeOPM].masculino > 0;

    }


    if (sexo === "F") {

        return vagas[nomeOPM].feminino > 0;

    }


    return false;

}


// ======================================================
// 9. CONSUMIR VAGA
// ======================================================

function consumirVaga(vagas, opm, sexo) {

    if (sexo === "M") {

        vagas[opm].masculino--;

    }


    if (sexo === "F") {

        vagas[opm].feminino--;

    }

}


// ======================================================
// 10. REALIZAR ESCOLHA
// ======================================================

function realizarEscolha(
    candidatos,
    opcoes,
    vagas
) {


    const resultado = [];


    candidatos.forEach(candidato => {


        const escolha =
            opcoes[candidato.rg];


        let opmAlocada = "";


        if (escolha) {


            const possibilidades = [

                escolha.opcao1,

                escolha.opcao2,

                escolha.opcao3

            ];


            // Tenta 1ª, 2ª e 3ª opção

            for (
                const opm of possibilidades
            ) {


                if (
                    possuiVaga(
                        vagas,
                        opm,
                        candidato.sexo
                    )
                ) {


                    opmAlocada =
                        normalizarTexto(opm);


                    consumirVaga(
                        vagas,
                        opmAlocada,
                        candidato.sexo
                    );


                    break;

                }

            }

        }


        resultado.push({

            classificacao:
                candidato.classificacao,

            rg:
                candidato.rg,

            nome:
                candidato.nome,

            sexo:
                candidato.sexo,

            opcaoAlocada:
                opmAlocada || "SEM ALOCAÇÃO"

        });

    });


    return resultado;

}


// ======================================================
// 11. MOSTRAR RESULTADO NA TELA
// ======================================================

function mostrarResultado(resultado) {

    const tabela =
        document.getElementById(
            "resultado"
        );


    tabela.innerHTML = "";


    let alocados = 0;
    let pendentes = 0;


    resultado.forEach(item => {


        const linha =
            document.createElement("tr");


        const classe =
            item.opcaoAlocada ===
            "SEM ALOCAÇÃO"
                ? "sem-alocacao"
                : "";


        linha.innerHTML = `

            <td>
                ${item.classificacao}
            </td>

            <td>
                ${item.rg}
            </td>

            <td>
                ${item.nome}
            </td>

            <td>
                ${item.sexo}
            </td>

            <td class="${classe}">
                ${item.opcaoAlocada}
            </td>

        `;


        tabela.appendChild(linha);


        if (
            item.opcaoAlocada ===
            "SEM ALOCAÇÃO"
        ) {

            pendentes++;

        } else {

            alocados++;

        }

    });


    document.getElementById(
        "total-candidatos"
    ).textContent = resultado.length;


    document.getElementById(
        "total-alocados"
    ).textContent = alocados;


    document.getElementById(
        "total-pendentes"
    ).textContent = pendentes;

}


// ======================================================
// 12. BOTÃO PROCESSAR
// ======================================================

const botao =
    document.getElementById(
        "process-button"
    );


botao.addEventListener(
    "click",
    async function () {


        try {


            const arquivoClassificacao =
                document.getElementById(
                    "classificacao"
                ).files[0];


            const arquivoOpcoes =
                document.getElementById(
                    "opcoes"
                ).files[0];


            const arquivoVagas =
                document.getElementById(
                    "vagas"
                ).files[0];


            if (
                !arquivoClassificacao ||
                !arquivoOpcoes ||
                !arquivoVagas
            ) {


                alert(
                    "Anexe as planilhas de Classificação, Opções e Vagas antes de processar."
                );


                return;

            }


            botao.textContent =
                "⏳ Processando...";


            botao.disabled = true;


            // Ler arquivos

            const dadosClassificacao =
                await lerExcel(
                    arquivoClassificacao
                );


            const dadosOpcoes =
                await lerExcel(
                    arquivoOpcoes
                );


            const dadosVagas =
                await lerExcel(
                    arquivoVagas
                );


            // Processar

            const candidatos =
                processarClassificacao(
                    dadosClassificacao
                );


            const opcoes =
                processarOpcoes(
                    dadosOpcoes
                );


            const vagas =
                processarVagas(
                    dadosVagas
                );


            // Realizar escolha

            const resultado =
                realizarEscolha(
                    candidatos,
                    opcoes,
                    vagas
                );


            // Mostrar resultado

            mostrarResultado(
                resultado
            );


            console.log(
                "Candidatos:",
                candidatos
            );


            console.log(
                "Opções:",
                opcoes
            );


            console.log(
                "Vagas:",
                vagas
            );


            console.log(
                "Resultado:",
                resultado
            );


            botao.textContent =
                "✓ Processamento concluído";


        } catch (erro) {


            console.error(erro);


            alert(
                "Erro no processamento:\n\n" +
                erro.message
            );


            botao.textContent =
                "⚙️ Processar escolha";


        } finally {


            botao.disabled = false;

        }

    }
);