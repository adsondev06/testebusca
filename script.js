// let dadosPlanilha = []; // Array para armazenar os dados da planilha
// let historicoBuscas = []; // Array para armazenar o histórico de buscas

// // Função para ler a planilha
// document.getElementById("input-excel").addEventListener("change", (event) => {
//     const file = event.target.files[0];

//     if (!file) {
//         alert("Por favor, selecione um arquivo.");
//         return;
//     }

//     const reader = new FileReader();

//     reader.onload = (e) => {
//         const data = new Uint8Array(e.target.result);
//         const workbook = XLSX.read(data, { type: 'array' });

//         // Supondo que a planilha está na primeira aba
//         const worksheet = workbook.Sheets[workbook.SheetNames[0]];
//         const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//         // Armazenar os dados da planilha no formato de objetos
//         dadosPlanilha = jsonData.map(row => ({
//             codigo: row[0], // Coluna A
//             driver: row[1]  // Coluna B
//         }));

//         console.log("Dados importados da planilha:", dadosPlanilha); // Para depuração
//     };

//     reader.onerror = () => {
//         alert("Erro ao ler o arquivo. Verifique se o formato está correto.");
//     };

//     reader.readAsArrayBuffer(file);
// });

// // Função para buscar pelo código
// function buscarPorCodigo() {
//     const codigoInput = document.getElementById("codigo").value.trim();
//     const resultadosDiv = document.getElementById("resultado");
//     resultadosDiv.innerHTML = ""; // Limpa resultados anteriores

//     // Valida entrada do usuário
//     if (codigoInput === "") {
//         resultadosDiv.innerHTML = "Por favor, insira um código para buscar.";
//         return;
//     }

//     // Busca pelo código
//     const resultado = dadosPlanilha.find(item => item.codigo.toString() === codigoInput);

//     if (resultado) {
//         resultadosDiv.innerHTML = `Código: <strong>${resultado.codigo}</strong> - Driver: <span class="highlight">${resultado.driver}</span>`;

//         // Adiciona ao histórico de buscas no início
//         historicoBuscas.unshift({
//             codigo: resultado.codigo,
//             driver: resultado.driver
//         });
//         atualizarHistorico();

//     } else {
//         resultadosDiv.innerHTML = "Nenhum resultado encontrado.";
//     }

//     // Limpa o campo de busca apenas quando o resultado for exibido ou não encontrado
//     document.getElementById("codigo").value = "";
//     document.getElementById("codigo").focus();
// }

// // Função para atualizar o histórico de buscas na tela
// function atualizarHistorico() {
//     const listaHistorico = document.getElementById("lista-historico");
//     listaHistorico.innerHTML = ""; // Limpa o histórico anterior

//     historicoBuscas.forEach(item => {
//         const li = document.createElement("li");
//         li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
//         listaHistorico.appendChild(li);
//     });
// }

// // Função para filtrar o histórico de buscas e ordenar numericamente
// function filtrarHistorico() {
//     const filtro = document.getElementById("filtro").value.toLowerCase();
//     const listaHistorico = document.getElementById("lista-historico");
//     listaHistorico.innerHTML = ""; // Limpa a lista atual

//     // Filtra os registros que correspondem ao filtro
//     const historicoFiltrado = historicoBuscas.filter(item => {
//         return item.codigo.toString().toLowerCase().includes(filtro) ||
//                item.driver.toLowerCase().includes(filtro);
//     });

//     // Ordena numericamente com base no número após "Rizzy"
//     historicoFiltrado.sort((a, b) => {
//         const numeroA = parseInt(a.driver.match(/\d+/)?.[0] || 0, 10); // Extrai o número de "Rizzy X"
//         const numeroB = parseInt(b.driver.match(/\d+/)?.[0] || 0, 10);
//         return numeroA - numeroB; // Compara os números para ordenar
//     });

//     // Exibe os registros filtrados e ordenados
//     historicoFiltrado.forEach(item => {
//         const li = document.createElement("li");
//         li.innerHTML = `Código: <strong>${item.codigo}</strong> - Driver: <span class="highlight">${item.driver}</span>`;
//         listaHistorico.appendChild(li);
//     });
// }

// // Função para gerar arquivo XLSX do histórico
// function gerarXLSX() {
//     const workbook = XLSX.utils.book_new();
//     const filtro = document.getElementById("filtro").value.toLowerCase();

//     // Se houver filtro, exporta apenas os itens filtrados
//     const historicoParaExportar = filtro 
//         ? historicoBuscas.filter(item => {
//             return item.codigo.toString().toLowerCase().includes(filtro) ||
//                    item.driver.toLowerCase().includes(filtro);
//         }).sort((a, b) => {
//             const numeroA = parseInt(a.driver.match(/\d+/)?.[0] || 0, 10);
//             const numeroB = parseInt(b.driver.match(/\d+/)?.[0] || 0, 10);
//             return numeroA - numeroB;
//         })
//         : historicoBuscas; // Caso contrário, exporta tudo

//     // Adiciona os dados do histórico na planilha
//     const worksheetData = [
//         ["Código", "Driver"], // Cabeçalhos das colunas
//         ...historicoParaExportar.map(item => [item.codigo, item.driver]) // Dados do histórico
//     ];

//     const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
//     XLSX.utils.book_append_sheet(workbook, worksheet, "Histórico de Buscas");
    
//     // Gera o arquivo e dispara o download
//     XLSX.writeFile(workbook, "historico_buscas.xlsx");
// }

// // Adiciona evento de entrada no campo de entrada de código com Enter
// document.getElementById("codigo").addEventListener("keydown", (event) => {
//     if (event.key === 'Enter') {
//         buscarPorCodigo(); // Executa a busca apenas ao pressionar Enter
//     }
// });

// // Alerta de confirmação ao recarregar a página
// window.addEventListener('beforeunload', function (event) {
//     const message = "Deseja recarregar a página? Seu histórico será perdido.";
//     event.returnValue = message; // Para a maioria dos navegadores
//     return message; // Para Firefox
// });



let dadosPlanilha = []; // Array para armazenar os dados da planilha
let historicoBuscas = []; // Array para armazenar o histórico de buscas
let impressoras = []; // Array para armazenar as impressoras disponíveis
let impressoraSelecionada = null; // Impressora selecionada pelo usuário

// Função para ler a planilha
document.getElementById("input-excel").addEventListener("change", (event) => {
    const file = event.target.files[0];

    if (!file) {
        alert("Por favor, selecione um arquivo.");
        return;
    }

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

    reader.onerror = () => {
        alert("Erro ao ler o arquivo. Verifique se o formato está correto.");
    };

    reader.readAsArrayBuffer(file);
});

// Função para buscar impressoras disponíveis
function buscarImpressoras() {
    ZebraBrowserPrint.getLocalPrinters()
        .then((printers) => {
            impressoras = printers;
            const select = document.getElementById("select-impressora");
            select.innerHTML = ""; // Limpa opções anteriores

            impressoras.forEach(printer => {
                const option = document.createElement("option");
                option.value = printer.name;
                option.textContent = printer.name;
                select.appendChild(option);
            });
        })
        .catch((error) => {
            console.error("Erro ao buscar impressoras:", error);
        });
}

// Função para buscar pelo código
function buscarPorCodigo() {
    const codigoInput = document.getElementById("codigo").value.trim();
    const resultadosDiv = document.getElementById("resultado");
    resultadosDiv.innerHTML = ""; // Limpa resultados anteriores

    // Valida entrada do usuário
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

        // Imprime automaticamente após mostrar o resultado
        if (impressoraSelecionada) {
            imprimirEtiqueta(resultado);
        } else {
        }
    } else {
        resultadosDiv.innerHTML = "Nenhum resultado encontrado.";
    }

    // Limpa o campo de busca apenas quando o resultado for exibido ou não encontrado
    document.getElementById("codigo").value = "";
    document.getElementById("codigo").focus();
}

// Função para imprimir a etiqueta
function imprimirEtiqueta(resultado) {
    const zpl = `
^XA
^FO50,100^ADN,36,20^FD${resultado.driver}^FS
^XZ
`;

    const selectedPrinter = impressoras.find(printer => printer.name === impressoraSelecionada);
    
    if (selectedPrinter) {
        const printer = new ZebraBrowserPrint.Printer(selectedPrinter.name);
        
        printer.print(zpl)
            .then(() => {
                console.log("Impressão enviada com sucesso.");
            })
            .catch((error) => {
                console.error("Erro ao enviar impressão:", error);
            });
    } else {
        console.error("Impressora não encontrada.");
    }
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
        const numeroA = parseInt(a.driver.match(/\d+/)?.[0] || 0, 10); // Extrai o número de "Rizzy X"
        const numeroB = parseInt(b.driver.match(/\d+/)?.[0] || 0, 10);
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
            const numeroA = parseInt(a.driver.match(/\d+/)?.[0] || 0, 10);
            const numeroB = parseInt(b.driver.match(/\d+/)?.[0] || 0, 10);
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

// Evento para selecionar impressora
document.getElementById("select-impressora").addEventListener("change", (event) => {
    impressoraSelecionada = event.target.value;
});

// Adiciona evento de entrada no campo de entrada de código com Enter
document.getElementById("codigo").addEventListener("keydown", (event) => {
    if (event.key === 'Enter') {
        buscarPorCodigo(); // Executa a busca apenas ao pressionar Enter
    }
});

// Alerta de confirmação ao recarregar a página
window.addEventListener('beforeunload', function (event) {
    const message = "Deseja recarregar a página? Seu histórico será perdido.";
    event.returnValue = message; // Para a maioria dos navegadores
    return message; // Para Firefox
});

// Chama a função para buscar impressoras quando a página carregar
document.addEventListener("DOMContentLoaded", buscarImpressoras);
