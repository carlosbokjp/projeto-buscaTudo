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

// Elementos XML
const xmlFileInput = document.getElementById('xmlFileInput');
const btnProcessarXml = document.getElementById('btnProcessarXml');
const statusXml = document.getElementById('statusXml');
const resultadosXml = document.getElementById('resultadosXml');
const tabelaProdutosXml = document.getElementById('tabelaProdutosXml');
const btnExportarXmlCSV = document.getElementById('btnExportarXmlCSV');

console.log('✅ Script carregado!');

// ========== FUNÇÃO PARA VERIFICAR SE NCM TEM 8 DÍGITOS ==========
function ncmTem8Digitos(codigo) {
    if (!codigo) return false;
    const numeros = codigo.replace(/\D/g, '');
    return numeros.length === 8;
}

// ========== FUNÇÃO PARA FORMATAR CAMPOS ==========

function formatarCampoCNPJ(valor, tipo) {
    if (!valor || valor === 'null' || valor === 'undefined' || valor === '') return 'Não informado';
    valor = String(valor).trim();
    
    switch(tipo) {
        case 'telefone':
            // Remove tudo que não é número
            const numeros = valor.replace(/\D/g, '');
            
            if (numeros.length === 10) {
                return numeros.replace(/^(\d{2})(\d{4})(\d{4})$/, '($1) $2-$3');
            }
            if (numeros.length === 11) {
                return numeros.replace(/^(\d{2})(\d{5})(\d{4})$/, '($1) $2-$3');
            }
            if (numeros.length === 8) {
                return numeros.replace(/^(\d{4})(\d{4})$/, '$1-$2');
            }
            if (numeros.length === 9) {
                return numeros.replace(/^(\d{5})(\d{4})$/, '$1-$2');
            }
            // Se não conseguir formatar, retorna o valor original
            return valor;
            
        case 'cep':
            const cepNumeros = valor.replace(/\D/g, '');
            if (cepNumeros.length === 8) {
                return cepNumeros.replace(/^(\d{5})(\d{3})$/, '$1-$2');
            }
            return valor;
            
        case 'capitalSocial':
            const num = parseFloat(valor);
            return isNaN(num) ? 'Não informado' : new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(num);
            
        default:
            return valor;
    }
}

// ========== FUNÇÃO PARA IMPRIMIR CARTÃO CNPJ ==========
function imprimirCartaoCNPJ() {
    const dados = window.ultimoCnpjConsultado;
    if (!dados) {
        alert('Nenhum dado de CNPJ disponível para impressão!');
        return;
    }

    // Formatar dados para impressão
    const cidadeNome = dados.cidade?.nome || 'Não informado';
    const uf = dados.uf || 'Não informado';
    const porteDescricao = dados.porte?.descricao || 'Não informado';
    const temSocios = dados.qsa && dados.qsa.length > 0;
    
    // Telefones
    let telefone1 = 'Não informado';
    let telefone2 = 'Não informado';
    
    if (dados.telefone1) {
        telefone1 = formatarCampoCNPJ(dados.telefone1, 'telefone');
    } else if (dados.ddd_telefone_1) {
        telefone1 = `(${dados.ddd_telefone_1}) ${dados.telefone_1}`;
    } else if (dados.telefone) {
        telefone1 = formatarCampoCNPJ(dados.telefone, 'telefone');
    }
    
    if (dados.telefone2) {
        telefone2 = formatarCampoCNPJ(dados.telefone2, 'telefone');
    } else if (dados.ddd_telefone_2) {
        telefone2 = `(${dados.ddd_telefone_2}) ${dados.telefone_2}`;
    }
    
    // Email
    let email = dados.email || 'Não informado';
    if (dados.email === null || dados.email === undefined || dados.email === '') {
        email = 'Não informado';
    }

    // Criar janela de impressão
    const janelaImpressao = window.open('', '_blank');
    
    // HTML do cartão
    const htmlCartao = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>Cartão CNPJ - ${dados.cnpj || ''}</title>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    padding: 0;
                    background: #fff;
                }
                .cartao {
                    max-width: 800px;
                    margin: 0 auto;
                    border: 2px solid #2563eb;
                    border-radius: 15px;
                    padding: 25px;
                    background: white;
                    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
                }
                .cabecalho {
                    text-align: center;
                    border-bottom: 2px solid #2563eb;
                    padding-bottom: 15px;
                    margin-bottom: 20px;
                }
                .cabecalho h1 {
                    color: #2563eb;
                    margin: 0;
                    font-size: 24px;
                    text-transform: uppercase;
                }
                .cabecalho .data {
                    color: #666;
                    font-size: 12px;
                    margin-top: 5px;
                }
                .secao {
                    margin-bottom: 20px;
                    page-break-inside: avoid;
                }
                .secao h3 {
                    color: #2563eb;
                    border-bottom: 1px solid #2563eb;
                    padding-bottom: 5px;
                    margin: 10px 0;
                    font-size: 16px;
                }
                .grid-2 {
                    display: grid;
                    grid-template-columns: repeat(2, 1fr);
                    gap: 15px;
                }
                .campo {
                    margin-bottom: 10px;
                }
                .campo .label {
                    font-weight: bold;
                    color: #666;
                    font-size: 12px;
                    text-transform: uppercase;
                }
                .campo .valor {
                    font-size: 14px;
                    color: #333;
                    margin-top: 2px;
                }
                .destaque {
                    background: #f0f4ff;
                    padding: 15px;
                    border-radius: 8px;
                    margin: 10px 0;
                }
                .socio-item {
                    border-left: 3px solid #2563eb;
                    padding-left: 10px;
                    margin-bottom: 10px;
                }
                .rodape {
                    margin-top: 30px;
                    text-align: center;
                    font-size: 11px;
                    color: #999;
                    border-top: 1px dashed #ccc;
                    padding-top: 10px;
                }
                @media print {
                    body { margin: 0; }
                    .cartao { box-shadow: none; border-color: #000; }
                    .cabecalho h1 { color: #000; }
                    .secao h3 { color: #000; border-bottom-color: #000; }
                }
            </style>
        </head>
        <body>
            <div class="cartao">
                <div class="cabecalho">
                    <h1>📋 CARTÃO CNPJ</h1>
                    <div class="data">Documento gerado em: ${new Date().toLocaleString('pt-BR')}</div>
                </div>
                
                <div class="secao">
                    <h3>🏢 DADOS PRINCIPAIS</h3>
                    <div class="grid-2">
                        <div class="campo">
                            <div class="label">CNPJ</div>
                            <div class="valor">${dados.cnpj || 'Não informado'}</div>
                        </div>
                        <div class="campo">
                            <div class="label">Situação</div>
                            <div class="valor">${dados.descricao_situacao_cadastral || 'Não informado'}</div>
                        </div>
                    </div>
                    <div class="campo">
                        <div class="label">Razão Social</div>
                        <div class="valor">${dados.razao_social || 'Não informado'}</div>
                    </div>
                    <div class="campo">
                        <div class="label">Nome Fantasia</div>
                        <div class="valor">${dados.nome_fantasia || 'Não informado'}</div>
                    </div>
                </div>

                <div class="secao">
                    <h3>📅 INFORMAÇÕES CADASTRAIS</h3>
                    <div class="grid-2">
                        <div class="campo">
                            <div class="label">Abertura</div>
                            <div class="valor">${dados.data_inicio_atividade ? new Date(dados.data_inicio_atividade).toLocaleDateString('pt-BR') : 'Não informado'}</div>
                        </div>
                        <div class="campo">
                            <div class="label">Capital Social</div>
                            <div class="valor">${dados.capital_social ? formatarCampoCNPJ(dados.capital_social, 'capitalSocial') : 'Não informado'}</div>
                        </div>
                        <div class="campo">
                            <div class="label">Porte</div>
                            <div class="valor">${porteDescricao}</div>
                        </div>
                    </div>
                </div>

                <div class="secao">
                    <h3>📍 ENDEREÇO</h3>
                    <div class="campo">
                        <div class="valor">
                            ${dados.logradouro || ''}, ${dados.numero || 'S/N'} ${dados.complemento || ''}<br>
                            ${dados.bairro || ''} - ${cidadeNome}/${uf}<br>
                            <strong>CEP:</strong> ${dados.cep ? formatarCampoCNPJ(dados.cep, 'cep') : 'Não informado'}
                        </div>
                    </div>
                </div>

                <div class="secao">
                    <h3>📞 CONTATO</h3>
                    <div class="grid-2">
                        <div class="campo">
                            <div class="label">Telefone 1</div>
                            <div class="valor">${telefone1}</div>
                        </div>
                        <div class="campo">
                            <div class="label">Telefone 2</div>
                            <div class="valor">${telefone2}</div>
                        </div>
                        <div class="campo">
                            <div class="label">Email</div>
                            <div class="valor">${email}</div>
                        </div>
                    </div>
                </div>

                <div class="secao">
                    <h3>🏢 ATIVIDADE PRINCIPAL</h3>
                    <div class="campo">
                        <div class="label">CNAE</div>
                        <div class="valor">${dados.cnae_fiscal || 'Não informado'}</div>
                    </div>
                    <div class="campo">
                        <div class="label">Descrição</div>
                        <div class="valor">${dados.cnae_fiscal_descricao || 'Não informado'}</div>
                    </div>
                </div>

                
                <div class="rodape">
                    Sistema de Consulta CNPJ - Documento gerado em ${new Date().toLocaleString('pt-BR')}
                </div>
            </div>
            <script>
                window.onload = function() { window.print(); }
            <\/script>
        </body>
        </html>
    `;
    
    janelaImpressao.document.write(htmlCartao);
    janelaImpressao.document.close();
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
    } else if (nomeAba === 'xml') {
        document.querySelectorAll('.tab-button')[3].classList.add('active');
        document.getElementById('abaXml').classList.add('active');
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
    // Armazena os dados globalmente para impressão
    window.ultimoCnpjConsultado = dados;
    
    // Formatar datas
    const dataAbertura = dados.data_inicio_atividade ? new Date(dados.data_inicio_atividade).toLocaleDateString('pt-BR') : 'Não informado';
    const dataSituacao = dados.data_situacao_cadastral ? new Date(dados.data_situacao_cadastral).toLocaleDateString('pt-BR') : 'Não informado';
    
    // Telefones
    let telefone1 = 'Não informado';
    let telefone2 = 'Não informado';
    
    if (dados.ddd_telefone_1) {
        telefone1 = formatarCampoCNPJ(dados.ddd_telefone_1, 'telefone');
    }
    
    if (dados.ddd_telefone_2) {
        telefone2 = formatarCampoCNPJ(dados.ddd_telefone_2, 'telefone');
    }
    
    // Email (veio como null neste exemplo)
    const email = dados.email || 'Não informado';
    
    // Formatar regime tributário
    const regimeTributario = dados.regime_tributario || [];
    const ultimoRegime = regimeTributario.length > 0 ? regimeTributario[regimeTributario.length - 1] : null;
    
    return `
        <div class="cnpj-info">
            <!-- DADOS PRINCIPAIS -->
            <div class="cnpj-section">
                <h4>📋 DADOS PRINCIPAIS</h4>
                <p><strong>CNPJ:</strong> ${dados.cnpj || 'Não informado'}</p>
                <p><strong>Razão Social:</strong> ${dados.razao_social || 'Não informado'}</p>
                <p><strong>Nome Fantasia:</strong> ${dados.nome_fantasia || 'Não informado'}</p>
                <p><strong>Natureza Jurídica:</strong> ${dados.natureza_juridica || 'Não informado'} (${dados.codigo_natureza_juridica || ''})</p>
                <p><strong>Porte:</strong> ${dados.porte || 'Não informado'} (Código: ${dados.codigo_porte || ''})</p>
                <p><strong>Capital Social:</strong> ${dados.capital_social ? formatarCampoCNPJ(dados.capital_social, 'capitalSocial') : 'Não informado'}</p>
            </div>
            
            <!-- SITUAÇÃO CADASTRAL -->
            <div class="cnpj-section">
                <h4>⚖️ SITUAÇÃO CADASTRAL</h4>
                <p><strong>Situação:</strong> ${dados.descricao_situacao_cadastral || 'Não informado'} (Código: ${dados.situacao_cadastral || ''})</p>
                <p><strong>Data da Situação:</strong> ${dataSituacao}</p>
                <p><strong>Motivo:</strong> ${dados.descricao_motivo_situacao_cadastral || 'Não informado'} (Código: ${dados.motivo_situacao_cadastral || ''})</p>
                <p><strong>Situação Especial:</strong> ${dados.situacao_especial || 'Não informado'}</p>
                <p><strong>Data Situação Especial:</strong> ${dados.data_situacao_especial ? new Date(dados.data_situacao_especial).toLocaleDateString('pt-BR') : 'Não informado'}</p>
            </div>
            
            <!-- ENDEREÇO -->
            <div class="cnpj-section">
                <h4>📍 ENDEREÇO</h4>
                <p><strong>Tipo Logradouro:</strong> ${dados.descricao_tipo_de_logradouro || ''}</p>
                <p><strong>Logradouro:</strong> ${dados.logradouro || ''}, ${dados.numero || 'S/N'}</p>
                <p><strong>Complemento:</strong> ${dados.complemento || 'Não informado'}</p>
                <p><strong>Bairro:</strong> ${dados.bairro || 'Não informado'}</p>
                <p><strong>Município:</strong> ${dados.municipio || 'Não informado'} (IBGE: ${dados.codigo_municipio_ibge || dados.codigo_municipio || ''})</p>
                <p><strong>UF:</strong> ${dados.uf || 'Não informado'}</p>
                <p><strong>CEP:</strong> ${dados.cep ? formatarCampoCNPJ(dados.cep, 'cep') : 'Não informado'}</p>
                <p><strong>País:</strong> ${dados.pais || 'Brasil'}</p>
            </div>
            
            <!-- CONTATO -->
            <div class="cnpj-section">
                <h4>📞 CONTATO</h4>
                <p><strong>Telefone 1:</strong> ${telefone1}</p>
                <p><strong>Telefone 2:</strong> ${telefone2}</p>
                <p><strong>Fax:</strong> ${dados.ddd_fax ? formatarCampoCNPJ(dados.ddd_fax, 'telefone') : 'Não informado'}</p>
                <p><strong>Email:</strong> ${email}</p>
            </div>
            
            <!-- ATIVIDADES ECONÔMICAS -->
            <div class="cnpj-section">
                <h4>🏭 ATIVIDADES ECONÔMICAS</h4>
                <p><strong>CNAE Fiscal:</strong> ${dados.cnae_fiscal || 'Não informado'}</p>
                <p><strong>Descrição CNAE:</strong> ${dados.cnae_fiscal_descricao || 'Não informado'}</p>
                
                ${dados.cnaes_secundarios && dados.cnaes_secundarios.length > 0 && dados.cnaes_secundarios[0].codigo !== 0 ? `
                    <p><strong>CNAEs Secundários:</strong></p>
                    <ul style="margin-top: 5px;">
                        ${dados.cnaes_secundarios.map(cnae => `
                            <li>Código: ${cnae.codigo} - ${cnae.descricao}</li>
                        `).join('')}
                    </ul>
                ` : '<p><strong>CNAEs Secundários:</strong> Não informado</p>'}
            </div>
            
            <!-- REGIME TRIBUTÁRIO -->
            <div class="cnpj-section">
                <h4>💰 REGIME TRIBUTÁRIO</h4>
                <p><strong>Opção pelo Simples:</strong> ${dados.opcao_pelo_simples ? 'Sim' : 'Não'}</p>
                <p><strong>Data Opção Simples:</strong> ${dados.data_opcao_pelo_simples ? new Date(dados.data_opcao_pelo_simples).toLocaleDateString('pt-BR') : 'Não informado'}</p>
                <p><strong>Data Exclusão Simples:</strong> ${dados.data_exclusao_do_simples ? new Date(dados.data_exclusao_do_simples).toLocaleDateString('pt-BR') : 'Não informado'}</p>
                
                <p><strong>Opção pelo MEI:</strong> ${dados.opcao_pelo_mei ? 'Sim' : 'Não'}</p>
                <p><strong>Data Opção MEI:</strong> ${dados.data_opcao_pelo_mei ? new Date(dados.data_opcao_pelo_mei).toLocaleDateString('pt-BR') : 'Não informado'}</p>
                <p><strong>Data Exclusão MEI:</strong> ${dados.data_exclusao_do_mei ? new Date(dados.data_exclusao_do_mei).toLocaleDateString('pt-BR') : 'Não informado'}</p>
                
                ${ultimoRegime ? `
                    <p><strong>Último Regime:</strong> ${ultimoRegime.forma_de_tributacao} (${ultimoRegime.ano})</p>
                ` : ''}
            </div>
            
            <!-- QUADRO SOCIETÁRIO -->
            <div class="cnpj-section">
                <h4>👥 QUADRO SOCIETÁRIO</h4>
                ${dados.qsa && dados.qsa.length > 0 ? dados.qsa.map(socio => `
                    <div class="socio-item">
                        <p><strong>${socio.nome_socio || 'Nome não informado'}</strong></p>
                        <p><strong>Qualificação:</strong> ${socio.qualificacao_socio || 'Não informado'} (Código: ${socio.codigo_qualificacao_socio || ''})</p>
                        <p><strong>CPF/CNPJ:</strong> ${socio.cnpj_cpf_do_socio || 'Não informado'}</p>
                        <p><strong>Data Entrada:</strong> ${socio.data_entrada_sociedade ? new Date(socio.data_entrada_sociedade).toLocaleDateString('pt-BR') : 'Não informado'}</p>
                        <p><strong>Faixa Etária:</strong> ${socio.faixa_etaria || 'Não informado'}</p>
                        <p><strong>Representante Legal:</strong> ${socio.nome_representante_legal || 'Não informado'} (CPF: ${socio.cpf_representante_legal || 'Não informado'})</p>
                        <p><strong>Qualificação Representante:</strong> ${socio.qualificacao_representante_legal || 'Não informado'}</p>
                    </div>
                `).join('') : '<p>Não informado</p>'}
            </div>
            
            <!-- INFORMAÇÕES DA MATRIZ/FILIAL -->
            <div class="cnpj-section">
                <h4>🏢 INFORMAÇÕES DA EMPRESA</h4>
                <p><strong>Tipo:</strong> ${dados.descricao_identificador_matriz_filial || 'Não informado'} (Código: ${dados.identificador_matriz_filial || ''})</p>
                <p><strong>Data de Abertura:</strong> ${dataAbertura}</p>
                <p><strong>Ente Federativo:</strong> ${dados.ente_federativo_responsavel || 'Não informado'}</p>
                <p><strong>Nome no Exterior:</strong> ${dados.nome_cidade_no_exterior || 'Não informado'}</p>
            </div>
            
            <div style="margin-top: 20px; text-align: center; display: flex; gap: 10px; justify-content: center;">
                <button onclick="imprimirCartaoCNPJ()" style="background: #2563eb; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px;">
                    🖨️ Imprimir Cartão CNPJ
                </button>
                
            </div>
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
        <div style="overflow-x: auto;">
            <table style="width:100%; border-collapse: collapse; min-width: 600px;">
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
        </div>
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

// ========== FUNÇÕES PARA XML NFE ==========
// Função para extrair produtos do XML
function extrairProdutosXML(xmlString) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "text/xml");
    
    // Namespace da NFe
    const ns = {
        nfe: "http://www.portalfiscal.inf.br/nfe"
    };
    
    // Função para obter valor de tag com namespace
    function getTagValue(parent, tagName) {
        const element = parent.getElementsByTagNameNS(ns.nfe, tagName)[0];
        return element ? element.textContent : '';
    }
    
    // Encontrar a tag infNFe (pode estar em níveis diferentes)
    let infNFe = xmlDoc.getElementsByTagNameNS(ns.nfe, "infNFe")[0];
    
    // Se não encontrou com namespace, tenta sem namespace
    if (!infNFe) {
        infNFe = xmlDoc.getElementsByTagName("infNFe")[0];
    }
    
    if (!infNFe) {
        throw new Error("Estrutura de NFe não encontrada no XML");
    }
    
    // Encontrar todos os produtos (det)
    const dets = infNFe.getElementsByTagNameNS(ns.nfe, "det");
    const listaProdutos = [];
    
    // Se não encontrou com namespace, tenta sem namespace
    const produtos = dets.length > 0 ? dets : infNFe.getElementsByTagName("det");
    
    for (let i = 0; i < produtos.length; i++) {
        // Tenta encontrar prod com namespace
        let prod = produtos[i].getElementsByTagNameNS(ns.nfe, "prod")[0];
        
        // Se não encontrou, tenta sem namespace
        if (!prod) {
            prod = produtos[i].getElementsByTagName("prod")[0];
        }
        
        if (prod) {
            const produto = {
                codigo: getTagValue(prod, "cProd"),
                nome: getTagValue(prod, "xProd"),
                ncm: getTagValue(prod, "NCM"),
                cfop: getTagValue(prod, "CFOP"),
                unidade: getTagValue(prod, "uCom"),
                quantidade: getTagValue(prod, "qCom"),
                valorUnitario: getTagValue(prod, "vUnCom"),
                valorTotal: getTagValue(prod, "vProd"),
                ean: getTagValue(prod, "cEAN"),
                eanTrib: getTagValue(prod, "cEANTrib")
            };
            
            listaProdutos.push(produto);
        }
    }
    
    return listaProdutos;
}

// Função para exibir tabela de produtos do XML
function exibirTabelaProdutosXml(produtos) {
    window.ultimosProdutosXml = produtos;
    
    let html = `
        <div style="margin-bottom: 15px; padding: 10px; background: #f0f4ff; border-radius: 8px;">
            <strong>Total de produtos:</strong> ${produtos.length}
        </div>
        <div style="overflow-x: auto;">
            <table style="width:100%; border-collapse: collapse; min-width: 1200px;">
                <thead>
                    <tr style="background: #2563eb; color: white;">
                        <th style="padding: 12px; text-align: left;">Código</th>
                        <th style="padding: 12px; text-align: left;">Produto</th>
                        <th style="padding: 12px; text-align: left;">NCM</th>
                        <th style="padding: 12px; text-align: left;">CFOP</th>
                        <th style="padding: 12px; text-align: left;">GTIN Comercial (cEAN)</th>
                        <th style="padding: 12px; text-align: left;">GTIN Tributável (cEANTrib)</th>
                        <th style="padding: 12px; text-align: right;">Qtd</th>
                        <th style="padding: 12px; text-align: left;">UN</th>
                        <th style="padding: 12px; text-align: right;">Valor Unit.</th>
                        <th style="padding: 12px; text-align: right;">Total</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    produtos.forEach(p => {
        // Formatar valores para exibição
        const quantidade = p.quantidade ? parseFloat(p.quantidade).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0,00';
        const valorUnit = p.valorUnitario ? parseFloat(p.valorUnitario).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0,00';
        const valorTotal = p.valorTotal ? parseFloat(p.valorTotal).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0,00';
        
        html += `
            <tr style="border-bottom: 1px solid #e5e7eb;">
                <td style="padding: 10px;">${p.codigo || '---'}</td>
                <td style="padding: 10px;">${p.nome || '---'}</td>
                <td style="padding: 10px; font-weight: bold; color: #2563eb;">${p.ncm || '---'}</td>
                <td style="padding: 10px;">${p.cfop || '---'}</td>
                <td style="padding: 10px; font-family: monospace;">${p.ean || '---'}</td>
                <td style="padding: 10px; font-family: monospace; color: #f59e0b;">${p.eanTrib || '---'}</td>
                <td style="padding: 10px; text-align: right;">${quantidade}</td>
                <td style="padding: 10px;">${p.unidade || '---'}</td>
                <td style="padding: 10px; text-align: right;">R$ ${valorUnit}</td>
                <td style="padding: 10px; text-align: right;">R$ ${valorTotal}</td>
            </tr>
        `;
    });
    
    html += `
                </tbody>
            </table>
        </div>
    `;
    
    tabelaProdutosXml.innerHTML = html;
    resultadosXml.classList.remove('hidden');
}

// Função para processar XML
btnProcessarXml.addEventListener('click', async () => {
    if (!xmlFileInput.files || xmlFileInput.files.length === 0) {
        statusXml.innerText = "❌ Selecione um ou mais arquivos XML!";
        statusXml.style.color = "red";
        return;
    }
    
    statusXml.innerText = `📂 Processando ${xmlFileInput.files.length} arquivo(s)...`;
    statusXml.style.color = "#2563eb";
    resultadosXml.classList.add('hidden');
    
    const todosProdutos = [];
    const erros = [];
    
    try {
        for (let i = 0; i < xmlFileInput.files.length; i++) {
            const file = xmlFileInput.files[i];
            try {
                const text = await file.text();
                const produtos = extrairProdutosXML(text);
                
                // Adiciona informação de qual arquivo veio cada produto
                produtos.forEach(p => {
                    p.arquivoOrigem = file.name;
                });
                
                todosProdutos.push(...produtos);
            } catch (e) {
                erros.push(`${file.name}: ${e.message}`);
            }
        }
        
        if (todosProdutos.length === 0) {
            statusXml.innerText = "❌ Nenhum produto encontrado nos XMLs!";
            statusXml.style.color = "red";
            return;
        }
        
        exibirTabelaProdutosXml(todosProdutos);
        
        let mensagem = `✅ Processamento concluído! ${todosProdutos.length} produtos encontrados.`;
        if (erros.length > 0) {
            mensagem += ` ⚠️ ${erros.length} arquivo(s) com erro.`;
        }
        statusXml.innerText = mensagem;
        statusXml.style.color = erros.length > 0 ? "orange" : "green";
        
    } catch (error) {
        statusXml.innerText = `❌ Erro ao processar XML: ${error.message}`;
        statusXml.style.color = "red";
        console.error(error);
    }
});

// Função para exportar produtos do XML para CSV
function exportarProdutosXmlCSV() {
    if (!window.ultimosProdutosXml || window.ultimosProdutosXml.length === 0) {
        alert('❌ Não há produtos para exportar!');
        return;
    }
    
    const produtos = window.ultimosProdutosXml;
    const cabecalho = [
        'Arquivo Origem',
        'Código',
        'Produto',
        'NCM',
        'CFOP',
        'GTIN Comercial (cEAN)',
        'GTIN Tributável (cEANTrib)',
        'Quantidade',
        'Unidade',
        'Valor Unitário',
        'Valor Total'
    ];
    
    const linhas = produtos.map(p => [
        `"${p.arquivoOrigem || 'NFe'}"`,
        `"${p.codigo || ''}"`,
        `"${(p.nome || '').replace(/"/g, '""')}"`,
        p.ncm || '',
        p.cfop || '',
        p.ean || '',
        p.eanTrib || '',
        p.quantidade ? parseFloat(p.quantidade).toFixed(2).replace('.', ',') : '0,00',
        p.unidade || '',
        p.valorUnitario ? parseFloat(p.valorUnitario).toFixed(2).replace('.', ',') : '0,00',
        p.valorTotal ? parseFloat(p.valorTotal).toFixed(2).replace('.', ',') : '0,00'
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
    const nomeArquivo = `produtos_nfe_${dataStr}_${horaStr}.csv`;
    
    link.href = url;
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    statusXml.innerText = `✅ Arquivo "${nomeArquivo}" gerado com sucesso!`;
    statusXml.style.color = "green";
}

// Event listener do botão exportar XML
if (btnExportarXmlCSV) {
    btnExportarXmlCSV.addEventListener('click', exportarProdutosXmlCSV);
}

// ========== INICIALIZAÇÃO ==========
window.addEventListener('load', () => {
    console.log('🚀 Aplicação inicializada');
    // Sincroniza a base automaticamente
    setTimeout(() => btnSincronizar.click(), 500);
});
