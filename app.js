const folderInput = document.getElementById('folderInput');
const btnProcessar = document.getElementById('btnProcessar');
const btnExportar = document.getElementById('btnExportar');
const alerta = document.getElementById('alerta');
const resumoArquivos = document.getElementById('resumoArquivos');
const relatorioFinal = document.getElementById('relatorioFinal');

let dadosConsolidados = [];

btnProcessar.addEventListener('click', async () => {
    limparTela();

    const files = Array.from(folderInput.files || []);

    if (!files.length) {
        exibirAlerta('Selecione uma pasta antes de processar.', 'warning');
        return;
    }

    const arquivosExcel = files.filter(file => {
        const nome = file.name.toLowerCase();
        return nome.endsWith('.xlsx') || nome.endsWith('.xls');
    });

    if (!arquivosExcel.length) {
        exibirAlerta('Nenhum arquivo Excel encontrado na pasta selecionada.', 'warning');
        return;
    }

    try {
        btnProcessar.disabled = true;
        btnProcessar.innerText = 'Processando...';

        for (const file of arquivosExcel) {
            const contexto = obterContextoDoArquivo(file);
            const linhas = await processarArquivoExcel(file, contexto);

            dadosConsolidados.push(...linhas);

            imprimirResumoArquivo({
                arquivo: file.name,
                caminho: file.webkitRelativePath,
                pasta: contexto.nomePasta,
                data: contexto.data,
                linhas: linhas.length,
            });
        }

        imprimirRelatorioConsolidado(dadosConsolidados);

        btnExportar.classList.remove('d-none');

        exibirAlerta(
            `Processamento concluído. Total de linhas: ${dadosConsolidados.length}`,
            'success'
        );
    } catch (error) {
        console.error(error);
        exibirAlerta('Erro ao processar os arquivos Excel.', 'danger');
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

    const worksheet = XLSX.utils.json_to_sheet(dadosConsolidados);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Consolidado');

    XLSX.writeFile(workbook, 'relatorio-consolidado.xlsx');
});

async function processarArquivoExcel(file, contexto) {
    const arrayBuffer = await file.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
        type: 'array',
    });

    const primeiraAba = workbook.SheetNames[0];
    const sheet = workbook.Sheets[primeiraAba];

    const linhas = XLSX.utils.sheet_to_json(sheet, {
        defval: '',
    });

    return linhas.map((linha, index) => {
        const duracaoMinutos = converterDuracaoParaMinutos(linha['Duração']);
        const status = duracaoMinutos >= 20 ? 'Presente' : 'Falta';

        return {
            Data: contexto.data,
            ...linha,
            Status: status,
            pastaOrigem: contexto.nomePasta,
            arquivoOrigem: file.name,
            numeroLinha: index + 2,
        };
    });
}

function obterContextoDoArquivo(file) {
    const caminho = file.webkitRelativePath || file.name;

    const partes = caminho.split('/');

    // Exemplo:
    // PastaRaiz/2026-02-27 20_30 Bruna Corte - Estatuto.../arquivo.xlsx
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

    // Exemplo: "1 h 19 min", "2 h", "1h 05 min"
    if (/\d+\s*h\b/.test(texto)) {
        const matchHoras = texto.match(/(\d+)\s*h\b/);
        const matchMinutos = texto.match(/(\d+)\s*min\b/);

        const horas = matchHoras ? Number(matchHoras[1]) : 0;
        const minutos = matchMinutos ? Number(matchMinutos[1]) : 0;

        return horas * 60 + minutos;
    }

    // Exemplo: "25 min", "25 minutos"
    const matchMinutosTexto = texto.match(/(\d+)\s*min/);
    if (matchMinutosTexto) {
        return Number(matchMinutosTexto[1]);
    }

    // Exemplo: 00:25:30 ou 25:30
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

function imprimirRelatorioConsolidado(dados) {
    if (!dados.length) {
        relatorioFinal.innerHTML = `
      <div class="alert alert-warning">
        Nenhum dado encontrado.
      </div>
    `;
        return;
    }

    const colunas = obterTodasAsColunas(dados);

    let html = `
    <p><strong>Total consolidado de linhas:</strong> ${dados.length}</p>

    <table class="table table-bordered table-striped table-sm align-middle">
      <thead class="table-dark">
        <tr>
          ${colunas.map(coluna => `<th>${coluna}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

    dados.forEach(linha => {
        html += `
      <tr>
        ${colunas.map(coluna => `<td>${linha[coluna] ?? ''}</td>`).join('')}
      </tr>
    `;
    });

    html += `
      </tbody>
    </table>
  `;

    relatorioFinal.innerHTML = html;
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
    resumoArquivos.innerHTML = '';
    relatorioFinal.innerHTML = '';
    alerta.className = 'alert d-none mt-4';
    alerta.innerHTML = '';
    btnExportar.classList.add('d-none');
}

function exibirAlerta(mensagem, tipo) {
    alerta.className = `alert alert-${tipo} mt-4`;
    alerta.innerHTML = mensagem;
}