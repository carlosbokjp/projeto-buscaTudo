let baseNCM = []; // Memória para os itens NCM

// ========== ELEMENTOS DOM ==========
// Elementos NCM
const btnSincronizar = document.getElementById('btnSincronizarBase');
const buscaNcmInput = document.getElementById('buscaNcmInput');
const statusNcm = document.getElementById('statusNcm');
const listaNcm = document.getElementById('listaResultadosNcm');
const resultadoNcm = document.getElementById('resultadoNcm');
const resCodigo = document.getElementById('resCodigo');
const resDescricao = document.getElementById('resDescricao');

// Elementos CNPJ
const cnpjInput = document.getElementById('cnpjInput');
const btnBuscarCnpj = document.getElementById('btnBuscarCnpj');
const statusCnpj = document.getElementById('statusCnpj');
const resultadoCnpj = document.getElementById('resultadoCnpj');
const dadosCnpj = document.getElementById('dadosCnpj');

// Elementos Importação CSV
const csvFileInput = document.getElementById('csvFileInput');
const btnImportarCSV = document.getElementById('btnImportarCSV');
const statusImportacao = document.getElementById('statusImportacao');
const resultadosImportacao = document.getElementById('resultadosImportacao');
const tabelaResultados = document.getElementById('tabelaResultados');
const btnExportarCSV = document.getElementById('btnExportarCSV');

console.log('✅ Script carregado!');

// ========== FUNÇÃO PARA VERIFICAR SE NCM TEM 8 DÍGITOS ==========
function ncmTem8Digitos(codigo) {
    if (!codigo) return false;
    const numeros = codigo.replace(/\D/g, '');
    return numeros.length === 8;
}

// ========== FUNÇÃO PARA FORMATAR CAMPOS ==========
function formatarCampoCNPJ(valor, tipo) {
    if (!valor || valor === 'null' || valor === 'undefined') return 'Não informado';
    valor = String(valor);
    
    switch(tipo) {
        case 'telefone':
            if (valor.length === 10) return valor.replace(/^(\d{2})(\d{4})(\d{4})$/, '($1) $2-$3');
            if (valor.length === 11) return valor.replace(/^(\d{2})(\d{5})(\d{4})$/, '($1) $2-$3');
            return valor;
        case 'cep':
            if (valor.length === 8) return valor.replace(/^(\d{5})(\d{3})$/, '$1-$2');
            return valor;
        case 'capitalSocial':
            const num = parseFloat(valor);
            return isNaN(num) ? 'Não informado' : new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(num);
        default:
            return valor;
    }
}

// ========== FUNÇÃO PARA ALTERNAR ABAS ==========
window.mostrarAba = function(nomeAba) {
    console.log('🔄 Trocando para aba:', nomeAba);
    
    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    if (nomeAba === 'ncm') {
        document.querySelectorAll('.tab-button')[0].classList.add('active');
        document.getElementById('abaNcm').classList.add('active');
    } else if (nomeAba === 'cnpj') {
        document.querySelectorAll('.tab-button')[1].classList.add('active');
        document.getElementById('abaCnpj').classList.add('active');
    } else if (nomeAba === 'importar') {
        document.querySelectorAll('.tab-button')[2].classList.add('active');
        document.getElementById('abaImportar').classList.add('active');
    }
};

// ========== FUNÇÕES NCM ==========
btnSincronizar.addEventListener('click', async () => {
    statusNcm.innerText = "Sincronizando base oficial... aguarde.";
    statusNcm.style.color = "#2563eb";
    
    try {
        const resposta = await fetch('https://brasilapi.com.br/api/ncm/v1');
        if (!resposta.ok) throw new Error();
        
        baseNCM = await resposta.json();
        const ncmValidos = baseNCM.filter(item => ncmTem8Digitos(item.codigo));
        
        statusNcm.innerText = `✅ Base sincronizada: ${ncmValidos.length} itens com 8 dígitos.`;
        statusNcm.style.color = "green";
        btnSincronizar.style.backgroundColor = "#16a34a";
        btnSincronizar.innerText = "Base Atualizada";
    } catch (erro) {
        statusNcm.innerText = "❌ Erro ao conectar com a Brasil API.";
        statusNcm.style.color = "red";
    }
});

buscaNcmInput.addEventListener('input', (e) => {
    const termo = e.target.value.toLowerCase();
    listaNcm.innerHTML = "";

    if (termo.length < 3) {
        if (baseNCM.length > 0) statusNcm.innerText = "Digite pelo menos 3 caracteres...";
        return;
    }

    const resultados = baseNCM.filter(item => 
        ncmTem8Digitos(item.codigo) && 
        (item.codigo.includes(termo) || item.descricao.toLowerCase().includes(termo))
    ).slice(0, 50);

    statusNcm.innerText = `${resultados.length} resultados para "${termo}"`;

    resultados.forEach(item => {
        const li = document.createElement('li');
        li.className = "resultado-item";

        const regex = new RegExp(`(${termo})`, 'gi');
        const descDestacada = item.descricao.replace(regex, '<mark>$1</mark>');
        const codDestacado = item.codigo.replace(regex, '<mark>$1</mark>');

        li.innerHTML = `
            <span class="ncm-code">${codDestacado}</span>
            <span class="ncm-desc">${descDestacada}</span>
        `;
        
        li.title = "Clique para ver detalhes";
        li.onclick = () => {
            resCodigo.innerText = item.codigo;
            resDescricao.innerText = item.descricao;
            resultadoNcm.classList.remove('hidden');
        };

        listaNcm.appendChild(li);
    });
});

// ========== FUNÇÕES CNPJ ==========
cnpjInput.addEventListener('input', (e) => {
    let value = e.target.value.replace(/\D/g, '');
    if (value.length > 14) value = value.slice(0, 14);
    
    if (value.length > 12) {
        value = value.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
    } else if (value.length > 8) {
        value = value.replace(/^(\d{2})(\d{3})(\d{3})(\d{0,4})/, '$1.$2.$3/$4');
    } else if (value.length > 5) {
        value = value.replace(/^(\d{2})(\d{3})(\d{0,3})/, '$1.$2.$3');
    } else if (value.length > 2) {
        value = value.replace(/^(\d{2})(\d{0,3})/, '$1.$2');
    }
    e.target.value = value;
});

function criarHtmlDadosCNPJ(dados) {
    const cidadeNome = dados.cidade?.nome || 'Não informado';
    const uf = dados.uf || 'Não informado';
    const porteDescricao = dados.porte?.descricao || 'Não informado';
    const temSocios = dados.qsa && dados.qsa.length > 0;
    
    return `
        <div class="cnpj-info">
            <div class="cnpj-section">
                <h4>📋 Dados Principais</h4>
                <p><strong>CNPJ:</strong> ${dados.cnpj || 'Não informado'}</p>
                <p><strong>Razão Social:</strong> ${dados.razao_social || 'Não informado'}</p>
                <p><strong>Nome Fantasia:</strong> ${dados.nome_fantasia || 'Não informado'}</p>
                <p><strong>Situação:</strong> ${dados.descricao_situacao_cadastral || 'Não informado'}</p>
                <p><strong>Abertura:</strong> ${dados.data_inicio_atividade ? new Date(dados.data_inicio_atividade).toLocaleDateString('pt-BR') : 'Não informado'}</p>
                <p><strong>Capital:</strong> ${dados.capital_social ? formatarCampoCNPJ(dados.capital_social, 'capitalSocial') : 'Não informado'}</p>
                <p><strong>Porte:</strong> ${porteDescricao}</p>
            </div>
            
            <div class="cnpj-section">
                <h4>📍 Endereço</h4>
                <p>${dados.logradouro || ''}, ${dados.numero || 'S/N'} ${dados.complemento || ''}</p>
                <p>${dados.bairro || ''} - ${cidadeNome}/${uf}</p>
                <p><strong>CEP:</strong> ${dados.cep ? formatarCampoCNPJ(dados.cep, 'cep') : 'Não informado'}</p>
            </div>
            
            <div class="cnpj-section">
                <h4>📞 Contato</h4>
                <p><strong>Telefone 1:</strong> ${dados.telefone1 ? formatarCampoCNPJ(dados.telefone1, 'telefone') : 'Não informado'}</p>
                <p><strong>Telefone 2:</strong> ${dados.telefone2 ? formatarCampoCNPJ(dados.telefone2, 'telefone') : 'Não informado'}</p>
                <p><strong>Email:</strong> ${dados.email || 'Não informado'}</p>
            </div>
            
            <div class="cnpj-section">
                <h4>🏢 Atividade Principal</h4>
                <p><strong>CNAE:</strong> ${dados.cnae_fiscal || 'Não informado'}</p>
                <p><strong>Descrição:</strong> ${dados.cnae_fiscal_descricao || 'Não informado'}</p>
            </div>
            
            ${temSocios ? `
                <div class="cnpj-section">
                    <h4>👥 Quadro Societário</h4>
                    ${dados.qsa.map(socio => `
                        <div class="socio-item">
                            <p><strong>${socio.nome || 'Nome não informado'}</strong></p>
                            <p>${socio.qualificacao || ''}</p>
                        </div>
                    `).join('')}
                </div>
            ` : ''}
        </div>
    `;
}

async function buscarCNPJ() {
    let cnpj = cnpjInput.value.replace(/\D/g, '');
    
    if (cnpj.length !== 14) {
        statusCnpj.innerText = "❌ CNPJ deve conter 14 números";
        statusCnpj.style.color = "red";
        return;
    }
    
    if (/^(\d)\1{13}$/.test(cnpj)) {
        statusCnpj.innerText = "❌ CNPJ inválido";
        statusCnpj.style.color = "red";
        return;
    }
    
    statusCnpj.innerText = "🔍 Consultando CNPJ...";
    statusCnpj.style.color = "#2563eb";
    resultadoCnpj.classList.add('hidden');
    
    try {
        const resposta = await fetch(`https://brasilapi.com.br/api/cnpj/v1/${cnpj}`);
        
        if (!resposta.ok) {
            if (resposta.status === 404) throw new Error("CNPJ não encontrado");
            if (resposta.status === 400) throw new Error("CNPJ inválido");
            throw new Error("Erro na consulta");
        }
        
        const dados = await resposta.json();
        dadosCnpj.innerHTML = criarHtmlDadosCNPJ(dados);
        resultadoCnpj.classList.remove('hidden');
        statusCnpj.innerText = "✅ Consulta realizada com sucesso!";
        statusCnpj.style.color = "green";
        
    } catch (erro) {
        statusCnpj.innerText = `❌ Erro: ${erro.message}`;
        statusCnpj.style.color = "red";
    }
}

btnBuscarCnpj.addEventListener('click', buscarCNPJ);
cnpjInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') buscarCNPJ();
});

// ========== FUNÇÕES DE IMPORTAÇÃO CSV ==========
function lerCSV(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (event) => {
            const lines = event.target.result.split('\n');
            const produtos = [];
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line) {
                    const fields = line.split(';');
                    const nomeProduto = fields[0].trim();
                    if (nomeProduto) {
                        produtos.push({ nome: nomeProduto });
                    }
                }
            }
            resolve(produtos);
        };
        
        reader.onerror = reject;
        reader.readAsText(file, 'UTF-8');
    });
}

function buscarNCMProduto(nomeProduto) {
    if (!baseNCM || baseNCM.length === 0) {
        return { encontrado: false, ncm: null, descricao: null };
    }
    
    const termo = nomeProduto.toLowerCase();
    const palavras = termo.split(' ').filter(p => p.length > 2);
    
    // Primeiro tenta busca exata
    let resultados = baseNCM.filter(item => 
        ncmTem8Digitos(item.codigo) && 
        item.descricao.toLowerCase().includes(termo)
    );
    
    // Se não encontrar, tenta com palavras individuais
    if (resultados.length === 0 && palavras.length > 0) {
        resultados = baseNCM.filter(item => {
            if (!ncmTem8Digitos(item.codigo)) return false;
            const descLower = item.descricao.toLowerCase();
            return palavras.some(palavra => descLower.includes(palavra));
        });
    }
    
    if (resultados.length > 0) {
        return {
            encontrado: true,
            ncm: resultados[0].codigo,
            descricao: resultados[0].descricao
        };
    }
    
    return { encontrado: false, ncm: null, descricao: null };
}

function exibirTabelaResultados(resultados) {
    const encontrados = resultados.filter(r => r.encontrado).length;
    const naoEncontrados = resultados.filter(r => !r.encontrado).length;
    
    window.ultimosResultados = resultados;
    
    let html = `
        <div style="margin-bottom: 15px; padding: 10px; background: #f0f4ff; border-radius: 8px;">
            <strong>Total:</strong> ${resultados.length} produtos | 
            <span style="color: green;">✅ Encontrados: ${encontrados}</span> | 
            <span style="color: red;">❌ Não encontrados: ${naoEncontrados}</span>
        </div>
        <table style="width:100%; border-collapse: collapse;">
            <thead>
                <tr style="background: #2563eb; color: white;">
                    <th style="padding: 12px; text-align: left;">Produto</th>
                    <th style="padding: 12px; text-align: left;">NCM (8 dígitos)</th>
                    <th style="padding: 12px; text-align: left;">Descrição NCM</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    resultados.forEach(r => {
        const corLinha = r.encontrado ? '' : 'background-color: #fff0f0;';
        
        html += `
            <tr style="border-bottom: 1px solid #e5e7eb; ${corLinha}">
                <td style="padding: 10px;">${r.nome}</td>
                <td style="padding: 10px; font-weight: bold; font-family: monospace; ${r.encontrado ? 'color: #2563eb;' : 'color: #999;'}">
                    ${r.ncm || '--------'}
                </td>
                <td style="padding: 10px; ${r.encontrado ? '' : 'color: #999; font-style: italic;'}">
                    ${r.descricao || 'Não encontrado'}
                </td>
            </tr>
        `;
    });
    
    html += `
            </tbody>
        </table>
    `;
    
    tabelaResultados.innerHTML = html;
    resultadosImportacao.classList.remove('hidden');
}

btnImportarCSV.addEventListener('click', async () => {
    if (!csvFileInput.files || csvFileInput.files.length === 0) {
        statusImportacao.innerText = "❌ Selecione um arquivo CSV primeiro!";
        statusImportacao.style.color = "red";
        return;
    }
    
    if (baseNCM.length === 0) {
        statusImportacao.innerText = "❌ Base NCM não sincronizada. Clique em 'Sincronizar Base' primeiro!";
        statusImportacao.style.color = "red";
        return;
    }
    
    const file = csvFileInput.files[0];
    statusImportacao.innerText = `📂 Processando arquivo: ${file.name}...`;
    statusImportacao.style.color = "#2563eb";
    
    try {
        const produtos = await lerCSV(file);
        
        if (produtos.length === 0) {
            statusImportacao.innerText = "❌ Nenhum produto encontrado no arquivo!";
            statusImportacao.style.color = "red";
            return;
        }
        
        statusImportacao.innerText = `🔍 Buscando NCMs para ${produtos.length} produtos...`;
        
        const resultados = produtos.map(produto => ({
            ...produto,
            ...buscarNCMProduto(produto.nome)
        }));
        
        exibirTabelaResultados(resultados);
        
        const encontrados = resultados.filter(r => r.encontrado).length;
        statusImportacao.innerText = `✅ Processamento concluído! ${encontrados} de ${resultados.length} produtos encontrados.`;
        statusImportacao.style.color = "green";
        
    } catch (error) {
        statusImportacao.innerText = `❌ Erro ao processar arquivo: ${error.message}`;
        statusImportacao.style.color = "red";
    }
});

// ========== FUNÇÃO EXPORTAR CSV ==========
function exportarResultadosCSV() {
    if (!window.ultimosResultados || window.ultimosResultados.length === 0) {
        alert('❌ Não há resultados para exportar!');
        return;
    }
    
    const resultados = window.ultimosResultados;
    const cabecalho = ['Produto', 'NCM (8 dígitos)', 'Descrição NCM', 'Status'];
    
    const linhas = resultados.map(r => [
        `"${r.nome.replace(/"/g, '""')}"`,
        r.ncm || 'NÃO ENCONTRADO',
        `"${(r.descricao || 'Não encontrado').replace(/"/g, '""')}"`,
        r.encontrado ? 'ENCONTRADO' : 'NÃO ENCONTRADO'
    ]);
    
    const conteudoCSV = [
        cabecalho.join(';'),
        ...linhas.map(linha => linha.join(';'))
    ].join('\n');
    
    const blob = new Blob(["\uFEFF" + conteudoCSV], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    const data = new Date();
    const dataStr = data.toISOString().slice(0,10).replace(/-/g, '');
    const horaStr = data.toTimeString().slice(0,8).replace(/:/g, '');
    const nomeArquivo = `ncm_resultados_${dataStr}_${horaStr}.csv`;
    
    link.href = url;
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    statusImportacao.innerText = `✅ Arquivo "${nomeArquivo}" gerado com sucesso!`;
    statusImportacao.style.color = "green";
}

// Event listener do botão exportar
if (btnExportarCSV) {
    btnExportarCSV.addEventListener('click', exportarResultadosCSV);
}

// ========== INICIALIZAÇÃO ==========
window.addEventListener('load', () => {
    console.log('🚀 Aplicação inicializada');
    // Sincroniza a base automaticamente
    setTimeout(() => btnSincronizar.click(), 500);
});