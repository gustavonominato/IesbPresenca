const folderInput = document.getElementById('folderInput');
const btnProcessar = document.getElementById('btnProcessar');
const btnExportar = document.getElementById('btnExportar');
const alerta = document.getElementById('alerta');
const resumoArquivos = document.getElementById('resumoArquivos');
const relatorioFinal = document.getElementById('relatorioFinal');

const blocoResumoArquivos = document.getElementById('blocoResumoArquivos');
const collapseResumoArquivos = document.getElementById('collapseResumoArquivos');
const iconeCollapseResumoArquivos = document.getElementById('iconeCollapseResumoArquivos');

let dadosConsolidados = [];
let dataSelecionada = null;
let graficoStatus = null;
let modoResumoAluno = false;
let modoConsolidadoAluno = false;

btnProcessar.addEventListener('click', async () => {
    limparTela();

    const files = Array.from(folderInput.files || []);

    if (!files.length) {
        exibirAlerta('Selecione uma pasta antes de processar.', 'warning');
        return;
    }

    const arquivosExcel = files.filter(file => {
        const nome = file.name.toLowerCase();

        return (
            !nome.startsWith('~$') &&
            (nome.endsWith('.xlsx') || nome.endsWith('.xls'))
        );
    });

    if (!arquivosExcel.length) {
        exibirAlerta('Nenhum arquivo Excel encontrado na pasta selecionada.', 'warning');
        return;
    }

    try {
        btnProcessar.disabled = true;
        btnProcessar.innerText = 'Processando...';

        console.log('Iniciando processamento dos arquivos Excel...');
        console.log('Total de arquivos selecionados:', files.length);
        console.log('Total de arquivos Excel válidos:', arquivosExcel.length);
        console.table(arquivosExcel.map(file => ({
            nome: file.name,
            caminho: file.webkitRelativePath,
            tamanhoBytes: file.size,
            tipo: file.type,
        })));

        const promessas = arquivosExcel.map(file => {
            const contexto = obterContextoDoArquivo(file);

            console.log('Processando arquivo:', {
                nome: file.name,
                caminho: file.webkitRelativePath,
                tamanhoBytes: file.size,
                tipo: file.type,
                contexto,
            });

            return processarArquivoExcel(file, contexto)
                .then(linhas => {
                    console.log('Arquivo processado com sucesso:', {
                        nome: file.name,
                        caminho: file.webkitRelativePath,
                        linhas: linhas.length,
                    });

                    dadosConsolidados.push(...linhas);

                    imprimirResumoArquivo({
                        arquivo: file.name,
                        caminho: file.webkitRelativePath,
                        pasta: contexto.nomePasta,
                        data: formatarData(contexto.data),
                        linhas: linhas.length,
                    });
                })
                .catch(error => {
                    console.error('Erro ao processar arquivo:', {
                        nome: file.name,
                        caminho: file.webkitRelativePath,
                        tamanhoBytes: file.size,
                        tipo: file.type,
                        contexto,
                        error,
                    });

                    throw new Error(
                        `Erro ao processar o arquivo "${file.webkitRelativePath || file.name}": ${obterMensagemErro(error)}`
                    );
                });
        });

        await Promise.all(promessas);

        imprimirRelatorioConsolidado(dadosConsolidados);

        blocoResumoArquivos.classList.remove('d-none');
        btnExportar.classList.remove('d-none');

        exibirAlerta(
            `Processamento concluído. Total de linhas: ${dadosConsolidados.length}`,
            'success'
        );
    } catch (error) {
        console.error('Erro geral no processamento:', error);
        console.error('Mensagem:', error?.message);
        console.error('Stack:', error?.stack);

        exibirAlerta(`Erro ao processar os arquivos Excel: ${obterMensagemErro(error)}`, 'danger');
    } finally {
        btnProcessar.disabled = false;
        btnProcessar.innerText = 'Processar';
    }
});

btnExportar.addEventListener('click', () => {
    if (!dadosConsolidados.length) {
        exibirAlerta('Nenhum dado consolidado para exportar.', 'warning');
        return;
    }

    const dadosParaExportar = ordenarPorDataENome(

        dataSelecionada

            ? dadosConsolidados.filter(linha => linha.DataOriginal === dataSelecionada)

            : dadosConsolidados

    );

    const dadosTratados = dadosParaExportar.map(removerColunasOcultasEInternas);

    const worksheet = XLSX.utils.json_to_sheet(dadosTratados);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, dataSelecionada ? 'Lista de Presença' : 'Consolidado');

    const nomeArquivo = dataSelecionada
        ? `lista-presenca-${dataSelecionada}.xlsx`
        : 'relatorio-consolidado.xlsx';

    XLSX.writeFile(workbook, nomeArquivo);
});

async function processarArquivoExcel(file, contexto) {
    console.log('Lendo arrayBuffer do arquivo:', file.webkitRelativePath || file.name);

    let arrayBuffer;

    try {
        arrayBuffer = await file.arrayBuffer();
    } catch (error) {
        console.error('Falha ao ler arrayBuffer:', {
            arquivo: file.webkitRelativePath || file.name,
            error,
        });
        throw new Error(`Falha ao ler o arquivo no navegador. ${obterMensagemErro(error)}`);
    }

    console.log('arrayBuffer lido com sucesso:', {
        arquivo: file.webkitRelativePath || file.name,
        bytes: arrayBuffer.byteLength,
    });

    let workbook;

    try {
        workbook = XLSX.read(arrayBuffer, {
            type: 'array',
        });
    } catch (error) {
        console.error('Falha ao interpretar Excel com XLSX.read:', {
            arquivo: file.webkitRelativePath || file.name,
            error,
        });
        throw new Error(`Falha ao interpretar o Excel. O arquivo pode estar corrompido ou em formato incompatível. ${obterMensagemErro(error)}`);
    }

    console.log('Workbook carregado:', {
        arquivo: file.webkitRelativePath || file.name,
        abas: workbook.SheetNames,
    });

    const primeiraAba = workbook.SheetNames[0];

    if (!primeiraAba) {
        throw new Error('O arquivo não possui abas para leitura.');
    }

    const sheet = workbook.Sheets[primeiraAba];

    if (!sheet) {
        throw new Error(`A primeira aba "${primeiraAba}" não pôde ser encontrada.`);
    }

    let linhas;

    try {
        linhas = XLSX.utils.sheet_to_json(sheet, {
            defval: '',
        });
    } catch (error) {
        console.error('Falha ao converter aba em JSON:', {
            arquivo: file.webkitRelativePath || file.name,
            aba: primeiraAba,
            error,
        });
        throw new Error(`Falha ao converter a aba "${primeiraAba}" para JSON. ${obterMensagemErro(error)}`);
    }

    console.log('Linhas extraídas:', {
        arquivo: file.webkitRelativePath || file.name,
        aba: primeiraAba,
        quantidade: linhas.length,
        primeiraLinha: linhas[0] || null,
    });

    return linhas.map((linha, index) => {
        const duracaoMinutos = converterDuracaoParaMinutos(linha['Duração']);
        const status = duracaoMinutos >= 20 ? 'Presente' : 'Falta';

        const linhaNormalizada = normalizarLinha(linha);

        return {
            DataOriginal: contexto.data,
            Data: formatarData(contexto.data),
            ...linhaNormalizada,
            Status: status,
            pastaOrigem: contexto.nomePasta,
            arquivoOrigem: file.name,
            numeroLinha: index + 2,
        };
    });
}

function obterMensagemErro(error) {
    if (!error) {
        return 'Erro desconhecido.';
    }

    if (error.message) {
        return error.message;
    }

    return String(error);
}

function normalizarLinha(linha) {
    const novaLinha = {};

    Object.keys(linha).forEach(coluna => {
        const nomeNormalizado = coluna === 'Enviar e-mail' ? 'e-mail' : coluna;
        novaLinha[nomeNormalizado] = linha[coluna];
    });

    return novaLinha;
}

function ordenarPorDataENome(dados) {
    return [...dados].sort((a, b) => {
        const dataA = a.DataOriginal || '';
        const dataB = b.DataOriginal || '';

        if (dataA !== dataB) {
            return dataA.localeCompare(dataB);
        }

        const nomeA = obterNomeAluno(a);
        const nomeB = obterNomeAluno(b);

        return nomeA.localeCompare(nomeB, 'pt-BR', {
            sensitivity: 'base',
        });
    });
}

function obterNomeAluno(linha) {
    return String(
        linha['Nome'] ||
        linha['Nome completo'] ||
        linha['Primeiro nome'] ||
        ''
    ).trim();
}

function montarResumoPorAluno(dados, totalAulas) {
    const mapa = {};

    dados.forEach(linha => {
        const nome = String(linha['Nome'] || '').trim();
        const sobrenome = String(linha['Sobrenome'] || '').trim();

        const chave = `${nome}|${sobrenome}`.toLowerCase();

        if (!mapa[chave]) {
            mapa[chave] = {
                nome,
                sobrenome,
                presenca: 0,
                falta: 0,
                participacoes: 0,
                total: totalAulas,
                percentual: '0.00',
            };
        }

        if (linha.Status === 'Presente') {
            mapa[chave].presenca++;
        }

        if (linha.Status === 'Falta') {
            mapa[chave].falta++;
        }

        mapa[chave].participacoes++;
    });

    Object.values(mapa).forEach(aluno => {
        aluno.percentual = totalAulas > 0
            ? ((aluno.presenca / totalAulas) * 100).toFixed(2)
            : '0.00';
    });

    return Object.values(mapa).sort((a, b) => {
        const nomeA = `${a.nome} ${a.sobrenome}`;
        const nomeB = `${b.nome} ${b.sobrenome}`;

        return nomeA.localeCompare(nomeB, 'pt-BR', {
            sensitivity: 'base',
        });
    });
}

function imprimirRelatorioConsolidado(dados) {
    if (!dados.length) {
        relatorioFinal.innerHTML = `
      <div class="alert alert-warning">
        Nenhum dado encontrado.
      </div>
    `;
        return;
    }

    const datas = obterDatasDisponiveis(dados);

    let html = `
    <div class="mb-3">
      <button 
        class="btn ${modoConsolidadoAluno ? 'btn-primary' : 'btn-outline-primary'} btn-sm me-2 mb-2"
        onclick="selecionarConsolidadoAluno()"
      >
        Consolidado por aluno
      </button>

      <button 
        class="btn ${modoResumoAluno ? 'btn-primary' : 'btn-outline-primary'} btn-sm me-2 mb-2"
        onclick="selecionarResumoAluno()"
      >
        Resumo por Aluno
      </button>

      <button 
        class="btn ${dataSelecionada === null && !modoResumoAluno && !modoConsolidadoAluno ? 'btn-primary' : 'btn-outline-primary'} btn-sm me-2 mb-2"
        onclick="selecionarData(null)"
      >
        Todos
      </button>

      ${datas.map(data => `
        <button 
          class="btn ${dataSelecionada === data && !modoResumoAluno && !modoConsolidadoAluno ? 'btn-primary' : 'btn-outline-primary'} btn-sm me-2 mb-2"
          onclick="selecionarData('${data}')"
        >
          ${formatarData(data)}
        </button>
      `).join('')}
    </div>
  `;

    if (modoConsolidadoAluno) {
        const datasAulas = obterDatasDisponiveis(dados);
        const consolidadoPorAluno = montarConsolidadoPorAluno(dados, datasAulas);

        html += `
      <h4 class="mb-3">Consolidado por aluno</h4>

      <p><strong>Total de alunos:</strong> ${consolidadoPorAluno.length}</p>
      <p><strong>Total de aulas ministradas:</strong> ${datasAulas.length}</p>

      <table class="table table-bordered table-striped table-sm align-middle">
        <thead class="table-dark">
        <tr>
          <th>Nome</th>
          <th>Sobrenome</th>
          ${datasAulas.map(data => `
            <th class="text-center coluna-data-vertical">
              <span>${formatarData(data)}</span>
            </th>
          `).join('')}
          <th>Faltas</th>
        </tr>
        </thead>
        <tbody>
          ${consolidadoPorAluno.map(aluno => `
            <tr>
              <td>${aluno.nome}</td>
              <td>${aluno.sobrenome}</td>
              ${datasAulas.map(data => {
            const presente = aluno.presencasPorData[data] === 1;

            return presente
                ? `<td class="text-center bg-success text-white fw-bold">✓</td>`
                : `<td class="text-center bg-danger text-white fw-bold">–</td>`;
        }).join('')}
              <td class="text-center fw-bold">
                ${calcularFaltasPonderadas(aluno, datasAulas)}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

        relatorioFinal.innerHTML = html;
        return;
    }

    if (modoResumoAluno) {
        const totalAulas = obterDatasDisponiveis(dados).length;
        const resumoPorAluno = montarResumoPorAluno(dados, totalAulas);

        html += `
      <h4 class="mb-3">Resumo por Aluno</h4>

      <p><strong>Total de alunos:</strong> ${resumoPorAluno.length}</p>
      <p><strong>Total de aulas ministradas:</strong> ${totalAulas}</p>

      <table class="table table-bordered table-striped table-sm align-middle">
        <thead class="table-dark">
          <tr>
            <th>Nome</th>
            <th>Sobrenome</th>
            <th>Presença</th>
            <th>Falta</th>
            <th>Participações</th>
            <th>Total</th>
            <th>Percentual</th>
          </tr>
        </thead>
        <tbody>
          ${resumoPorAluno.map(aluno => `
            <tr>
              <td>${aluno.nome}</td>
              <td>${aluno.sobrenome}</td>
              <td class="bg-success text-white fw-bold">${aluno.presenca}</td>
              <td class="bg-danger text-white fw-bold">${aluno.falta}</td>
              <td class="fw-bold">${aluno.participacoes}</td>
              <td class="fw-bold">${aluno.total}</td>
              <td class="fw-bold">${aluno.percentual}%</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

        relatorioFinal.innerHTML = html;
        return;
    }

    const dadosFiltrados = ordenarPorDataENome(
        dataSelecionada
            ? dados.filter(linha => linha.DataOriginal === dataSelecionada)
            : dados
    );

    const titulo = dataSelecionada
        ? `Lista de Presença em ${formatarData(dataSelecionada)}`
        : 'Relatório consolidado';

    const dadosTratados = dadosFiltrados.map(removerColunasOcultasEInternas);
    const colunas = obterTodasAsColunas(dadosTratados);

    html += `
    <h4 class="mb-3">${titulo}</h4>

    <p><strong>Total de linhas:</strong> ${dadosTratados.length}</p>

    <div class="card mb-4">
      <div class="card-body">
        <h5 class="card-title">Totalizador de Presença</h5>
        <canvas id="graficoStatus" style="max-height: 300px;"></canvas>
      </div>
    </div>

    <table class="table table-bordered table-striped table-sm align-middle">
      <thead class="table-dark">
        <tr>
          ${colunas.map(coluna => `<th>${coluna}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

    dadosTratados.forEach(linha => {
        html += `
      <tr>
        ${colunas.map(coluna => {
            const valor = linha[coluna] ?? '';

            if (coluna === 'Status' && valor === 'Presente') {
                return `<td class="bg-success text-white fw-bold">${valor}</td>`;
            }

            if (coluna === 'Status' && valor === 'Falta') {
                return `<td class="bg-danger text-white fw-bold">${valor}</td>`;
            }

            return `<td>${valor}</td>`;
        }).join('')}
      </tr>
    `;
    });

    html += `
      </tbody>
    </table>
  `;

    relatorioFinal.innerHTML = html;

    renderizarGraficoStatus(dadosTratados);
}

function selecionarData(data) {
    modoConsolidadoAluno = false;
    modoResumoAluno = false;
    dataSelecionada = data;
    imprimirRelatorioConsolidado(dadosConsolidados);
}

function selecionarResumoAluno() {
    modoConsolidadoAluno = false;
    modoResumoAluno = true;
    dataSelecionada = null;
    imprimirRelatorioConsolidado(dadosConsolidados);
}

function selecionarConsolidadoAluno() {
    modoConsolidadoAluno = true;
    modoResumoAluno = false;
    dataSelecionada = null;
    imprimirRelatorioConsolidado(dadosConsolidados);
}

function removerColunasOcultasEInternas(linha) {
    const novaLinha = { ...linha };

    delete novaLinha.DataOriginal;
    delete novaLinha.pastaOrigem;
    delete novaLinha.arquivoOrigem;
    delete novaLinha.numeroLinha;

    return novaLinha;
}

function obterDatasDisponiveis(dados) {
    return Array.from(
        new Set(
            dados
                .map(linha => linha.DataOriginal)
                .filter(Boolean)
        )
    ).sort();
}

function formatarData(dataIso) {
    if (!dataIso) return '';

    const partes = dataIso.split('-');

    if (partes.length !== 3) {
        return dataIso;
    }

    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function obterContextoDoArquivo(file) {
    const caminho = file.webkitRelativePath || file.name;
    const partes = caminho.split('/');

    const nomePasta = partes.length >= 2 ? partes[partes.length - 2] : '';
    const data = extrairDataDoNome(nomePasta);

    return {
        nomePasta,
        data,
    };
}

function extrairDataDoNome(nome) {
    const match = nome.match(/^(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : '';
}

function converterDuracaoParaMinutos(valor) {
    if (!valor) return 0;

    if (typeof valor === 'number') {
        return valor * 24 * 60;
    }

    const texto = String(valor).trim().toLowerCase();

    if (/\d+\s*h\b/.test(texto)) {
        const matchHoras = texto.match(/(\d+)\s*h\b/);
        const matchMinutos = texto.match(/(\d+)\s*min\b/);

        const horas = matchHoras ? Number(matchHoras[1]) : 0;
        const minutos = matchMinutos ? Number(matchMinutos[1]) : 0;

        return horas * 60 + minutos;
    }

    const matchMinutosTexto = texto.match(/(\d+)\s*min/);
    if (matchMinutosTexto) {
        return Number(matchMinutosTexto[1]);
    }

    const matchHorario = texto.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);

    if (matchHorario) {
        const parte1 = Number(matchHorario[1]);
        const parte2 = Number(matchHorario[2]);
        const parte3 = Number(matchHorario[3] || 0);

        if (matchHorario[3]) {
            return parte1 * 60 + parte2 + parte3 / 60;
        }

        return parte1 + parte2 / 60;
    }

    const matchNumero = texto.match(/(\d+)/);
    return matchNumero ? Number(matchNumero[1]) : 0;
}

function imprimirResumoArquivo(item) {
    resumoArquivos.innerHTML += `
    <div class="card mb-2">
      <div class="card-body py-3">
        <strong>Data:</strong> ${item.data || '-'}<br>
        <strong>Pasta:</strong> ${item.pasta || '-'}<br>
        <strong>Arquivo:</strong> ${item.arquivo}<br>
        <strong>Caminho:</strong> ${item.caminho}<br>
        <strong>Linhas:</strong> ${item.linhas}
      </div>
    </div>
  `;
}

function obterTodasAsColunas(dados) {
    const colunas = new Set();

    dados.forEach(linha => {
        Object.keys(linha).forEach(coluna => colunas.add(coluna));
    });

    return Array.from(colunas);
}

function limparTela() {
    dadosConsolidados = [];
    dataSelecionada = null;
    resumoArquivos.innerHTML = '';
    relatorioFinal.innerHTML = '';
    alerta.className = 'alert d-none mt-4';
    alerta.innerHTML = '';
    btnExportar.classList.add('d-none');

    blocoResumoArquivos.classList.add('d-none');
    iconeCollapseResumoArquivos.innerText = '▶';
    const collapseInstance = bootstrap.Collapse.getOrCreateInstance(collapseResumoArquivos, {
        toggle: false,
    });
    collapseInstance.hide();
}

function exibirAlerta(mensagem, tipo) {
    alerta.className = `alert alert-${tipo} mt-4`;
    alerta.innerHTML = mensagem;
}

function renderizarGraficoStatus(dados) {

    const totalPresente = dados.filter(linha => linha.Status === 'Presente').length;

    const totalFalta = dados.filter(linha => linha.Status === 'Falta').length;

    const canvas = document.getElementById('graficoStatus');

    if (!canvas) return;

    if (graficoStatus) {

        graficoStatus.destroy();

    }

    graficoStatus = new Chart(canvas, {

        type: 'doughnut',

        data: {

            labels: ['Presente', 'Falta'],

            datasets: [

                {

                    data: [totalPresente, totalFalta],

                    backgroundColor: ['#198754', '#dc3545'],

                },

            ],

        },

        options: {

            responsive: true,

            plugins: {

                legend: {

                    position: 'bottom',

                },

            },

        },

    });

}

function montarConsolidadoPorAluno(dados, datasAulas) {
    const mapa = {};

    dados.forEach(linha => {
        const nome = String(linha['Nome'] || '').trim();
        const sobrenome = String(linha['Sobrenome'] || '').trim();

        const chave = `${nome}|${sobrenome}`.toLowerCase();

        if (!mapa[chave]) {
            mapa[chave] = {
                nome,
                sobrenome,
                presencasPorData: {},
            };

            datasAulas.forEach(data => {
                mapa[chave].presencasPorData[data] = 0;
            });
        }

        if (linha.Status === 'Presente' && linha.DataOriginal) {
            mapa[chave].presencasPorData[linha.DataOriginal] = 1;
        }
    });

    return Object.values(mapa).sort((a, b) => {
        const nomeA = `${a.nome} ${a.sobrenome}`;
        const nomeB = `${b.nome} ${b.sobrenome}`;

        return nomeA.localeCompare(nomeB, 'pt-BR', {
            sensitivity: 'base',
        });
    });
}

collapseResumoArquivos.addEventListener('show.bs.collapse', () => {
    iconeCollapseResumoArquivos.innerText = '▼';
});

collapseResumoArquivos.addEventListener('hide.bs.collapse', () => {
    iconeCollapseResumoArquivos.innerText = '▶';
});

function calcularFaltasPonderadas(aluno, datasAulas) {
    const quantidadeFaltas = datasAulas.filter(data => aluno.presencasPorData[data] !== 1).length;
    return quantidadeFaltas * 1.5;
}