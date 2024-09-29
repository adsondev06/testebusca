let dadosPlanilha = []; // Array para armazenar os dados da planilha
let historicoBuscas = []; // Array para armazenar o histórico de buscas

// Função para ler a planilha
document.getElementById("input-excel").addEventListener("change", (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Supondo que a planilha está na primeira aba
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Armazenar os dados da planilha no formato de objetos
        dadosPlanilha = jsonData.map(row => ({
            codigo: row[0], // Coluna A
            driver: row[1]  // Coluna B
        }));

        console.log("Dados importados da planilha:", dadosPlanilha); // Para depuração
    };

    reader.readAsArrayBuffer(file);
});

// Função para buscar pelo código
function buscarPorCodigo() {
    const codigoInput = document.getElementById("codigo").value.trim();
    const resultadosDiv = document.getElementById("resultado");
    resultadosDiv.innerHTML = ""; // Limpa resultados anteriores

    if (codigoInput === "") {
        resultadosDiv.innerHTML = "Por favor, insira um código para buscar.";
        return;
    }

    // Busca pelo código
    const resultado = dadosPlanilha.find(item => item.codigo.toString() === codigoInput);

    if (resultado) {
        resultadosDiv.innerHTML = `Código: <strong>${resultado.codigo}</strong> - Driver: <span class="highlight">${resultado.driver}</span>`;

        // Adiciona ao histórico de buscas no início
        historicoBuscas.unshift({
            codigo: resultado.codigo,
            driver: resultado.driver
        });
        atualizarHistorico();

    } else {
        resultadosDiv.innerHTML = "Nenhum resultado encontrado.";
    }

    // Limpa o campo de busca e mantém o foco
    document.getElementById("codigo").value = "";
    document.getElementById("codigo").focus();
}

// Função para atualizar o histórico de buscas na tela
function atualizarHistorico() {
    const listaHistorico = document.getElementById("lista-historico");
    listaHistorico.innerHTML = ""; // Limpa o histórico anterior

    historicoBuscas.forEach(item => {
        const li = document.createElement("li");
        li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
        listaHistorico.appendChild(li);
    });
}

// Função para filtrar o histórico de buscas e ordenar numericamente
function filtrarHistorico() {
    const filtro = document.getElementById("filtro").value.toLowerCase();
    const listaHistorico = document.getElementById("lista-historico");
    listaHistorico.innerHTML = ""; // Limpa a lista atual

    // Filtra os registros que correspondem ao filtro
    const historicoFiltrado = historicoBuscas.filter(item => {
        return item.codigo.toString().toLowerCase().includes(filtro) ||
               item.driver.toLowerCase().includes(filtro);
    });

    // Ordena numericamente com base no número após "Rizzy"
    historicoFiltrado.sort((a, b) => {
        const numeroA = parseInt(a.driver.match(/\d+/)[0], 10); // Extrai o número de "Rizzy X"
        const numeroB = parseInt(b.driver.match(/\d+/)[0], 10);
        return numeroA - numeroB; // Compara os números para ordenar
    });

    // Exibe os registros filtrados e ordenados
    historicoFiltrado.forEach(item => {
        const li = document.createElement("li");
        li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
        listaHistorico.appendChild(li);
    });
}

// Função para gerar arquivo XLSX do histórico
function gerarXLSX() {
    const workbook = XLSX.utils.book_new();
    const filtro = document.getElementById("filtro").value.toLowerCase();

    // Se houver filtro, exporta apenas os itens filtrados
    const historicoParaExportar = filtro 
        ? historicoBuscas.filter(item => {
            return item.codigo.toString().toLowerCase().includes(filtro) ||
                   item.driver.toLowerCase().includes(filtro);
        }).sort((a, b) => {
            const numeroA = parseInt(a.driver.match(/\d+/)[0], 10);
            const numeroB = parseInt(b.driver.match(/\d+/)[0], 10);
            return numeroA - numeroB;
        })
        : historicoBuscas; // Caso contrário, exporta tudo

    // Adiciona os dados do histórico na planilha
    const worksheetData = [
        ["Código", "Driver"], // Cabeçalhos das colunas
        ...historicoParaExportar.map(item => [item.codigo, item.driver]) // Dados do histórico
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Histórico de Buscas");
    
    // Gera o arquivo e dispara o download
    XLSX.writeFile(workbook, "historico_buscas.xlsx");
}

// Adiciona evento de entrada no campo de entrada de código
document.getElementById("codigo").addEventListener("input", () => {
    buscarPorCodigo(); // Chama a função de busca sempre que o usuário digita
});

// Alerta de confirmação ao recarregar a página
window.addEventListener('beforeunload', function (event) {
    const message = "Deseja recarregar a página? Seu histórico será perdido.";
    event.returnValue = message; // Para a maioria dos navegadores
    return message; // Para Firefox
});