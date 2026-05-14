// Armazena as imagens carregadas
let imagensCarregadas = {};
let planilhaCarregada = false;

document.getElementById('imagensFile').addEventListener('change', function(event) {
    const files = event.target.files;
    imagensCarregadas = {};
    const statusImagens = document.getElementById('status-imagens');
    
    if (files.length === 0) {
        statusImagens.textContent = 'Nenhuma imagem selecionada';
        statusImagens.className = 'status-indicador status-pendente';
        return;
    }
    
    statusImagens.textContent = `Carregando ${files.length} imagens...`;
    statusImagens.className = 'status-indicador status-carregando';
    
    let carregadas = 0;
    
    Array.from(files).forEach(file => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            imagensCarregadas[file.name] = e.target.result;
            carregadas++;
            
            if (carregadas === files.length) {
                statusImagens.textContent = `✅ ${carregadas} imagens carregadas`;
                statusImagens.className = 'status-indicador status-sucesso';
                console.log('Imagens carregadas:', Object.keys(imagensCarregadas));
            }
        };
        
        reader.onerror = function() {
            console.error(`Erro ao carregar imagem: ${file.name}`);
            statusImagens.textContent = `❌ Erro ao carregar: ${file.name}`;
            statusImagens.className = 'status-indicador status-erro';
        };
        
        reader.readAsDataURL(file);
    });
});

document.getElementById('excelFile').addEventListener('change', function (event) {
    const file = event.target.files[0];
    const statusPlanilha = document.getElementById('status-planilha');
    
    if (!file) {
        statusPlanilha.textContent = 'Nenhuma planilha selecionada';
        statusPlanilha.className = 'status-indicador status-pendente';
        return;
    }

    statusPlanilha.textContent = 'Carregando planilha...';
    statusPlanilha.className = 'status-indicador status-carregando';

    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (jsonData.length === 0) {
                alert("A planilha está vazia.");
                statusPlanilha.textContent = '❌ Planilha vazia';
                statusPlanilha.className = 'status-indicador status-erro';
                return;
            }
            
            planilhaCarregada = true;
            statusPlanilha.textContent = `✅ Planilha carregada (${file.name})`;
            statusPlanilha.className = 'status-indicador status-sucesso';
            
            processarDadosExcel(jsonData);
            
        } catch (error) {
            console.error("Erro ao processar o arquivo Excel:", error);
            alert("Erro ao processar o arquivo Excel. Verifique o console para mais detalhes.");
            statusPlanilha.textContent = '❌ Erro ao processar planilha';
            statusPlanilha.className = 'status-indicador status-erro';
        }
    };
    
    reader.onerror = function() {
        alert("Erro ao ler o arquivo.");
        statusPlanilha.textContent = '❌ Erro ao ler arquivo';
        statusPlanilha.className = 'status-indicador status-erro';
    };
    
    reader.readAsArrayBuffer(file);
});

function atualizarStatus(mensagem, tipo) {
    const statusContainer = document.getElementById('status-container');
    statusContainer.innerHTML = `<div class="status-mensagem status-${tipo}">${mensagem}</div>`;
}

function processarDadosExcel(jsonData) {
    const cabecalhos = jsonData[0];
    
    const mapeamentoColunas = {
        numtombo: encontrarIndice(cabecalhos, ['numtombo']),
        family: encontrarIndice(cabecalhos, ['family']),
        genus: encontrarIndice(cabecalhos, ['genus']),
        sp1: encontrarIndice(cabecalhos, ['sp1']),
        author1: encontrarIndice(cabecalhos, ['author1']),
        detby: encontrarIndice(cabecalhos, ['detby']),
        detdd: encontrarIndice(cabecalhos, ['detdd']),
        detmm: encontrarIndice(cabecalhos, ['detmm']),
        detyy: encontrarIndice(cabecalhos, ['detyy']),
        collector: encontrarIndice(cabecalhos, ['collector']),
        number: encontrarIndice(cabecalhos, ['number']),
        projeto: encontrarIndice(cabecalhos, ['projeto'])
    };
    
    const colunasFaltantes = Object.entries(mapeamentoColunas)
        .filter(([_, indice]) => indice === -1)
        .map(([nome]) => nome);
    
    if (colunasFaltantes.length > 0) {
        alert(`Colunas não encontradas na planilha: ${colunasFaltantes.join(', ')}`);
        console.error("Colunas faltantes:", colunasFaltantes);
        return;
    }
    
    const dadosProcessados = [];
    let imagensNaoEncontradas = [];
    
    for (let i = 1; i < jsonData.length; i++) {
        const linha = jsonData[i];
        
        if (!linha.some(celula => celula !== '' && celula !== null && celula !== undefined)) {
            continue;
        }
        
        const numtombo = String(linha[mapeamentoColunas.numtombo] || '').trim();
        const family = String(linha[mapeamentoColunas.family] || '').trim();
        const genus = String(linha[mapeamentoColunas.genus] || '').trim();
        const sp1 = String(linha[mapeamentoColunas.sp1] || '').trim();
        const author1 = String(linha[mapeamentoColunas.author1] || '').trim();
        const detby = String(linha[mapeamentoColunas.detby] || '').trim();
        const detdd = String(linha[mapeamentoColunas.detdd] || '').trim();
        const detmm = String(linha[mapeamentoColunas.detmm] || '').trim();
        const detyy = String(linha[mapeamentoColunas.detyy] || '').trim();
        const collector = String(linha[mapeamentoColunas.collector] || '').trim();
        const number = String(linha[mapeamentoColunas.number] || '').trim();
        const projeto = String(linha[mapeamentoColunas.projeto] || '').trim();
        
        const evb = numtombo ? `EVB ${numtombo}` : '';
        const codigoBarras = numtombo ? `EVB00${numtombo}.png` : '';
        
        // Verifica se a imagem foi carregada
        if (codigoBarras && !imagensCarregadas[codigoBarras]) {
            imagensNaoEncontradas.push(codigoBarras);
        }
        
        dadosProcessados.push({
            evb,
            familia: family,
            genero: genus,
            epiteto: sp1,
            autor: author1,
            determinador: detby,
            dia: detdd,
            mes: detmm,
            ano: detyy,
            Coletor: collector,
            NumeroColeta: number,
            codigoBarras: codigoBarras,
            projeto: projeto
        });
    }
    
    if (dadosProcessados.length === 0) {
        alert("Nenhum dado válido encontrado na planilha.");
        return;
    }
    
    // Exibe aviso sobre imagens não encontradas
    if (imagensNaoEncontradas.length > 0) {
        console.warn('Imagens não encontradas:', imagensNaoEncontradas);
        atualizarStatus(`${imagensNaoEncontradas.length} imagens de código de barras não foram encontradas no upload.`, 'aviso');
    }
    
    console.log(`Processados ${dadosProcessados.length} registros.`);
    criarEtiquetas(dadosProcessados);
}

function encontrarIndice(cabecalhos, possiveisNomes) {
    for (let i = 0; i < cabecalhos.length; i++) {
        const cabecalho = String(cabecalhos[i]).trim().toLowerCase();
        for (const nome of possiveisNomes) {
            if (cabecalho === nome.toLowerCase()) {
                return i;
            }
        }
    }
    
    for (let i = 0; i < cabecalhos.length; i++) {
        const cabecalho = String(cabecalhos[i]).trim().toLowerCase();
        for (const nome of possiveisNomes) {
            if (cabecalho.includes(nome.toLowerCase()) || nome.toLowerCase().includes(cabecalho)) {
                console.warn(`Coluna "${cabecalhos[i]}" encontrada como correspondência parcial para "${nome}"`);
                return i;
            }
        }
    }
    
    return -1;
}

document.getElementById('gerar-pdf').addEventListener('click', function () {
    console.log("Botão 'Gerar PDF' clicado.");

    const { jsPDF } = window.jspdf;
    if (!jsPDF) {
        console.error("jsPDF não está carregado corretamente.");
        return;
    }

    const doc = new jsPDF('p', 'mm', 'a4');
    const etiquetas = document.querySelectorAll('.etiqueta');

    if (etiquetas.length === 0) {
        console.error("Nenhuma etiqueta encontrada para exportar.");
        return;
    }

    const larguraEtiqueta = 105;
    const alturaEtiqueta = 29.7;
    const margemX = 0;
    const margemY = 10;
    const espacamentoX = 0;
    const espacamentoY = 0;
    const maxEtiquetasPorLinha = 2;

    let x = margemX;
    let y = margemY;

    function capturarEtiqueta(etiqueta, callback) {
        html2canvas(etiqueta, {
            scale: 4
        }).then(canvas => {
            const imgData = canvas.toDataURL('image/png', 1.0);
            callback(imgData);
        }).catch(error => {
            console.error("Erro ao capturar etiqueta:", error);
        });
    }

    function processarEtiqueta(index) {
        if (index >= etiquetas.length) {
            console.log("Todas as etiquetas foram processadas. Salvando PDF...");
            doc.save('etiquetas.pdf');
            return;
        }

        const etiqueta = etiquetas[index];
        capturarEtiqueta(etiqueta, function (imgData) {
            doc.addImage(imgData, 'PNG', x, y, larguraEtiqueta, alturaEtiqueta);

            x += larguraEtiqueta + espacamentoX;

            if ((index + 1) % maxEtiquetasPorLinha === 0) {
                x = margemX;
                y += alturaEtiqueta + espacamentoY;
            }

            if (y + alturaEtiqueta > doc.internal.pageSize.getHeight() - margemY) {
                doc.addPage();
                y = margemY;
            }

            processarEtiqueta(index + 1);
        });
    }

    processarEtiqueta(0);
});

function criarEtiquetas(dados) {
    const container = document.getElementById('etiquetas-container');
    container.innerHTML = '';

    dados.forEach((registro, index) => {
        const { evb, familia, genero, epiteto, autor, determinador, dia, mes, ano, Coletor, NumeroColeta, codigoBarras, projeto } = registro;

        const etiqueta = document.createElement('div');
        etiqueta.className = 'etiqueta';

        const colunaEsquerda = document.createElement('div');
        colunaEsquerda.className = 'coluna-esquerda';
        colunaEsquerda.innerHTML = `
            <div class="evb">${evb || ''}</div>
            <div class="familia">${familia || ''}</div>
            <div>
                <span class="genero">${genero || ''}</span>
                <span class="epiteto">${epiteto || ''}</span>
                <span class="autor">${autor || ''}</span>
            </div>
            <div class="Coletor">${Coletor ? `Col.: ${Coletor}` : ''}</div>
            <span class="NumeroColeta">${NumeroColeta || ''}</span>
            <div class="determinador">${determinador ? `Det.: ${determinador}` : ''}</div>
            <span class="data">${[dia, mes, ano].filter(Boolean).join('/') || ''}</span>
            <div class="projeto">${projeto || ''}</div>
        `;

        const colunaDireita = document.createElement('div');
        colunaDireita.className = 'coluna-direita';
        
        // Usa a imagem carregada se existir, senão tenta o caminho local
        let srcImagem = '';
        if (codigoBarras) {
            if (imagensCarregadas[codigoBarras]) {
                srcImagem = imagensCarregadas[codigoBarras];
            } else {
                srcImagem = `codigos_barras/${codigoBarras}`;
            }
        }
        
        colunaDireita.innerHTML = `
            <img src="LogoEVB.png" alt="Logo" class="logo">
            ${codigoBarras ? `<img src="${srcImagem}" alt="Código de Barras" class="codigo-barras" onerror="console.error('Erro ao carregar: ${codigoBarras}')">` : ''}
        `;

        etiqueta.appendChild(colunaEsquerda);
        etiqueta.appendChild(colunaDireita);

        container.appendChild(etiqueta);
    });
}