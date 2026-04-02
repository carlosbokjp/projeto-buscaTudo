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

// ========== MAPEAMENTO DOS GRUPOS NCM PELOS 4 PRIMEIROS DÍGITOS ==========
const gruposNCM = {
    // Seção I - Animais vivos e produtos do reino animal
    '0101': '0101 - Cavalos, asininos e muares, vivos',
    '0102': '0102 - Animais vivos da espécie bovina',
    '0103': '0103 - Animais vivos da espécie suína',
    '0104': '0104 - Animais vivos das espécies ovina e caprina',
    '0105': '0105 - Aves vivas',
    '0106': '0106 - Outros animais vivos',
    '0201': '0201 - Carnes de animais da espécie bovina, frescas ou refrigeradas',
    '0202': '0202 - Carnes de animais da espécie bovina, congeladas',
    '0203': '0203 - Carnes de animais da espécie suína, frescas, refrigeradas ou congeladas',
    '0204': '0204 - Carnes de animais das espécies ovina ou caprina, frescas, refrigeradas ou congeladas',
    '0205': '0205 - Carnes de animais da espécie equina, frescas, refrigeradas ou congeladas',
    '0206': '0206 - Miudezas comestíveis de animais',
    '0207': '0207 - Carnes e miudezas comestíveis de aves',
    '0208': '0208 - Outras carnes e miudezas comestíveis, frescas, refrigeradas ou congeladas',
    '0209': '0209 - Toucinho e gordura de porco',
    '0210': '0210 - Carnes e miudezas comestíveis, salgadas ou em salmoura',
    '0301': '0301 - Peixes vivos',
    '0302': '0302 - Peixes frescos ou refrigerados',
    '0303': '0303 - Peixes congelados',
    '0304': '0304 - Filés e outras carnes de peixe',
    '0305': '0305 - Peixes secos, salgados ou em salmoura',
    '0306': '0306 - Crustáceos',
    '0307': '0307 - Moluscos',
    '0308': '0308 - Invertebrados aquáticos',
    '0401': '0401 - Leite e creme de leite',
    '0402': '0402 - Leite em pó, granulado ou outras formas sólidas',
    '0403': '0403 - Iogurte e leite fermentado',
    '0404': '0404 - Soro de leite',
    '0405': '0405 - Manteiga e outras matérias gordas do leite',
    '0406': '0406 - Queijos',
    '0407': '0407 - Ovos de aves, com casca, frescos',
    '0408': '0408 - Ovos sem casca e gemas de ovos',
    '0409': '0409 - Mel natural',
    '0410': '0410 - Produtos comestíveis de origem animal',
    
    // Seção II - Produtos do reino vegetal
    '0601': '0601 - Bolbos, tubérculos e raízes',
    '0602': '0602 - Plantas vivas',
    '0603': '0603 - Flores e botões de flores',
    '0604': '0604 - Folhagens e ramos',
    '0701': '0701 - Batatas, frescas ou refrigeradas',
    '0702': '0702 - Tomates, frescos ou refrigerados',
    '0703': '0703 - Cebolas, chalotas, alhos, alhos-porós',
    '0704': '0704 - Couves, couves-flores e brócolis',
    '0705': '0705 - Alfaces e chicórias',
    '0706': '0706 - Cenouras, nabos e beterrabas',
    '0707': '0707 - Pepinos e cogombros',
    '0708': '0708 - Leguminosas',
    '0709': '0709 - Outros legumes e hortaliças',
    '0710': '0710 - Legumes e hortaliças congelados',
    '0711': '0711 - Legumes e hortaliças conservados',
    '0712': '0712 - Legumes e hortaliças secos',
    '0713': '0713 - Leguminosas secas',
    '0714': '0714 - Mandioca, araruta e outras raízes',
    '0801': '0801 - Cocos, castanhas e nozes',
    '0802': '0802 - Outras frutas de casca rija',
    '0803': '0803 - Bananas',
    '0804': '0804 - Tâmaras, figos, abacaxis, abacates',
    '0805': '0805 - Frutas cítricas',
    '0806': '0806 - Uvas',
    '0807': '0807 - Melões e melancias',
    '0808': '0808 - Maçãs e peras',
    '0809': '0809 - Damascos, cerejas, pêssegos',
    '0810': '0810 - Outras frutas frescas',
    '0811': '0811 - Frutas congeladas',
    '0812': '0812 - Frutas conservadas',
    '0813': '0813 - Frutas secas',
    '0814': '0814 - Cascas de frutas cítricas',
    '0901': '0901 - Café',
    '0902': '0902 - Chá',
    '0903': '0903 - Mate',
    '0904': '0904 - Pimenta',
    '0905': '0905 - Baunilha',
    '0906': '0906 - Canela',
    '0907': '0907 - Cravo',
    '0908': '0908 - Noz-moscada',
    '0909': '0909 - Sementes de anis, badiana, funcho',
    '0910': '0910 - Gengibre, açafrão, tomilho',
    
    // Seção III - Gorduras e óleos
    '1501': '1501 - Banha de porco',
    '1502': '1502 - Gorduras de animais',
    '1503': '1503 - Estearina',
    '1504': '1504 - Gorduras e óleos de peixes',
    '1505': '1505 - Lã de gordura',
    '1506': '1506 - Outras gorduras e óleos animais',
    '1507': '1507 - Óleo de soja',
    '1508': '1508 - Óleo de amendoim',
    '1509': '1509 - Azeite de oliva',
    '1510': '1510 - Outros óleos de oliva',
    '1511': '1511 - Óleo de palma',
    '1512': '1512 - Óleo de girassol',
    '1513': '1513 - Óleo de coco',
    '1514': '1514 - Óleo de canola',
    '1515': '1515 - Outras gorduras e óleos vegetais',
    '1516': '1516 - Gorduras e óleos hidrogenados',
    '1517': '1517 - Margarina',
    '1518': '1518 - Gorduras e óleos modificados',
    
    // Seção IV - Produtos das indústrias alimentares
    '1601': '1601 - Enchidos',
    '1602': '1602 - Preparações de carnes',
    '1603': '1603 - Extratos de carnes',
    '1604': '1604 - Preparações de peixes',
    '1605': '1605 - Preparações de crustáceos e moluscos',
    '1701': '1701 - Açúcares',
    '1702': '1702 - Outros açúcares',
    '1703': '1703 - Melaços',
    '1704': '1704 - Produtos de confeitaria',
    '1801': '1801 - Cacau em grão',
    '1802': '1802 - Cacau em massa',
    '1803': '1803 - Pasta de cacau',
    '1804': '1804 - Manteiga de cacau',
    '1805': '1805 - Cacau em pó',
    '1806': '1806 - Chocolate e preparações alimentícias',
    '1901': '1901 - Extratos de malte',
    '1902': '1902 - Massas alimentícias',
    '1903': '1903 - Tapioca',
    '1904': '1904 - Cereais preparados',
    '1905': '1905 - Produtos de panificação',
    '2001': '2001 - Legumes e hortaliças preparados',
    '2002': '2002 - Tomates preparados',
    '2003': '2003 - Cogumelos preparados',
    '2004': '2004 - Legumes congelados preparados',
    '2005': '2005 - Legumes preparados',
    '2006': '2006 - Frutas cristalizadas',
    '2007': '2007 - Doces e geleias',
    '2008': '2008 - Frutas preparadas',
    '2009': '2009 - Sucos de frutas',
    '2101': '2101 - Extratos de café',
    '2102': '2102 - Fermentos',
    '2103': '2103 - Molhos e condimentos',
    '2104': '2104 - Sopas e caldos',
    '2105': '2105 - Sorvetes',
    '2106': '2106 - Preparações alimentícias diversas',
    '2201': '2201 - Águas',
    '2202': '2202 - Águas gaseificadas',
    '2203': '2203 - Cervejas',
    '2204': '2204 - Vinhos',
    '2205': '2205 - Vermutes',
    '2206': '2206 - Bebidas fermentadas',
    '2207': '2207 - Álcool etílico',
    '2208': '2208 - Bebidas alcoólicas',
    '2209': '2209 - Vinagre',
    
    // Seção V - Produtos minerais
    '2501': '2501 - Sal',
    '2502': '2502 - Pirites',
    '2503': '2503 - Enxofre',
    '2504': '2504 - Grafite',
    '2505': '2505 - Areias',
    '2506': '2506 - Quartzo',
    '2507': '2507 - Caulim',
    '2508': '2508 - Argilas',
    '2509': '2509 - Giz',
    '2510': '2510 - Fosfatos',
    '2511': '2511 - Sulfato de bário',
    '2512': '2512 - Terras diatomáceas',
    '2513': '2513 - Pedra-pomes',
    '2514': '2514 - Ardósia',
    '2515': '2515 - Mármore',
    '2516': '2516 - Granito',
    '2517': '2517 - Seixos',
    '2518': '2518 - Dolomita',
    '2519': '2519 - Carbonato de magnésio',
    '2520': '2520 - Gesso',
    '2521': '2521 - Calcário',
    '2522': '2522 - Cal',
    '2523': '2523 - Cimento',
    '2524': '2524 - Amianto',
    '2525': '2525 - Mica',
    '2526': '2526 - Esteatite',
    '2528': '2528 - Boratos',
    '2529': '2529 - Feldspato',
    '2530': '2530 - Matérias minerais diversas',
    
    // Seção VI - Produtos das indústrias químicas
    '2801': '2801 - Flúor, cloro, bromo, iodo',
    '2802': '2802 - Enxofre sublimado',
    '2803': '2803 - Carbono',
    '2804': '2804 - Hidrogênio e gases raros',
    '2805': '2805 - Metais alcalinos',
    '2806': '2806 - Ácido clorídrico',
    '2807': '2807 - Ácido sulfúrico',
    '2808': '2808 - Ácido nítrico',
    '2809': '2809 - Pentaóxido de fósforo',
    '2810': '2810 - Óxidos de boro',
    '2811': '2811 - Ácidos inorgânicos',
    '2812': '2812 - Haletos',
    '2813': '2813 - Sulfetos',
    '2814': '2814 - Amônia',
    '2815': '2815 - Hidróxido de sódio',
    '2816': '2816 - Hidróxido de magnésio',
    '2817': '2817 - Óxido de zinco',
    '2818': '2818 - Óxido de alumínio',
    '2819': '2819 - Óxido de cromo',
    '2820': '2820 - Óxido de manganês',
    '2821': '2821 - Óxido de ferro',
    '2822': '2822 - Óxido de cobalto',
    '2823': '2823 - Óxido de titânio',
    '2824': '2824 - Óxido de chumbo',
    '2825': '2825 - Hidrazina',
    '2826': '2826 - Fluoretos',
    '2827': '2827 - Cloretos',
    '2828': '2828 - Hipocloritos',
    '2829': '2829 - Cloratos',
    '2830': '2830 - Sulfetos',
    '2831': '2831 - Ditionitos',
    '2832': '2832 - Sulfitos',
    '2833': '2833 - Sulfatos',
    '2834': '2834 - Nitritos',
    '2835': '2835 - Fosfinatos',
    '2836': '2836 - Carbonatos',
    '2837': '2837 - Cianetos',
    '2839': '2839 - Silicatos',
    '2840': '2840 - Boratos',
    '2841': '2841 - Sais de oxoácidos',
    '2842': '2842 - Sais de ácidos inorgânicos',
    '2843': '2843 - Colóides',
    '2844': '2844 - Elementos químicos radioativos',
    '2845': '2845 - Isótopos',
    '2846': '2846 - Compostos de terras raras',
    '2847': '2847 - Água oxigenada',
    '2848': '2848 - Fosfetos',
    '2849': '2849 - Carbonetos',
    '2850': '2850 - Hidretos',
    '2851': '2851 - Água destilada',
    '2852': '2852 - Compostos de mercúrio',
    '2853': '2853 - Compostos inorgânicos',
    '2901': '2901 - Hidrocarbonetos',
    '2902': '2902 - Hidrocarbonetos cíclicos',
    '2903': '2903 - Derivados halogenados',
    '2904': '2904 - Derivados sulfonados',
    '2905': '2905 - Álcoois',
    '2906': '2906 - Álcoois cíclicos',
    '2907': '2907 - Fenóis',
    '2908': '2908 - Derivados halogenados de fenóis',
    '2909': '2909 - Éteres',
    '2910': '2910 - Epóxidos',
    '2911': '2911 - Acetais',
    '2912': '2912 - Aldeídos',
    '2913': '2913 - Derivados halogenados de aldeídos',
    '2914': '2914 - Cetonas',
    '2915': '2915 - Ácidos monocarboxílicos',
    '2916': '2916 - Ácidos carboxílicos cíclicos',
    '2917': '2917 - Ácidos policarboxílicos',
    '2918': '2918 - Ácidos carboxílicos com funções oxigenadas',
    '2919': '2919 - Ésteres fosfóricos',
    '2920': '2920 - Ésteres de ácidos inorgânicos',
    '2921': '2921 - Aminas',
    '2922': '2922 - Amino-álcoois',
    '2923': '2923 - Sais de amônio',
    '2924': '2924 - Amidas',
    '2925': '2925 - Compostos de função carboximida',
    '2926': '2926 - Nitrilas',
    '2927': '2927 - Compostos diazoicos',
    '2928': '2928 - Hidrazinas',
    '2929': '2929 - Compostos isocianatos',
    '2930': '2930 - Compostos organossulfurados',
    '2931': '2931 - Compostos organometálicos',
    '2932': '2932 - Compostos heterocíclicos',
    '2933': '2933 - Compostos heterocíclicos com nitrogênio',
    '2934': '2934 - Ácidos nucleicos',
    '2935': '2935 - Sulfonamidas',
    '2936': '2936 - Provitaminas',
    '2937': '2937 - Hormônios',
    '2938': '2938 - Glicosídeos',
    '2939': '2939 - Alcalóides',
    '2940': '2940 - Açúcares quimicamente puros',
    '2941': '2941 - Antibióticos',
    '2942': '2942 - Outros compostos orgânicos',
    '3001': '3001 - Glândulas',
    '3002': '3002 - Sangue humano e animal',
    '3003': '3003 - Medicamentos',
    '3004': '3004 - Medicamentos dosificados',
    '3005': '3005 - Artigos para curativos',
    '3006': '3006 - Produtos farmacêuticos',
    '3101': '3101 - Adubos animais ou vegetais',
    '3102': '3102 - Adubos minerais nitrogenados',
    '3103': '3103 - Adubos fosfatados',
    '3104': '3104 - Adubos potássicos',
    '3105': '3105 - Adubos mistos',
    '3201': '3201 - Taninos',
    '3202': '3202 - Corantes sintéticos',
    '3203': '3203 - Corantes vegetais',
    '3204': '3204 - Corantes orgânicos',
    '3205': '3205 - Lacas',
    '3206': '3206 - Pigmentos',
    '3207': '3207 - Pigmentos preparados',
    '3208': '3208 - Tintas e vernizes',
    '3209': '3209 - Tintas aquosas',
    '3210': '3210 - Tintas preparadas',
    '3211': '3211 - Secantes',
    '3212': '3212 - Pigmentos em pó',
    '3213': '3213 - Tintas para arte',
    '3214': '3214 - Massas de vidraceiro',
    '3215': '3215 - Tintas de impressão',
    '3301': '3301 - Óleos essenciais',
    '3302': '3302 - Misturas de odoríferos',
    '3303': '3303 - Perfumes',
    '3304': '3304 - Produtos de beleza',
    '3305': '3305 - Produtos capilares',
    '3306': '3306 - Produtos de higiene bucal',
    '3307': '3307 - Produtos de barbear',
    '3401': '3401 - Sabões',
    '3402': '3402 - Agentes de superfície',
    '3403': '3403 - Preparações lubrificantes',
    '3404': '3404 - Ceras artificiais',
    '3405': '3405 - Graxas',
    '3406': '3406 - Velas',
    '3407': '3407 - Pastas para modelar',
    '3501': '3501 - Caseína',
    '3502': '3502 - Albumina',
    '3503': '3503 - Gelatina',
    '3504': '3504 - Peptonas',
    '3505': '3505 - Dextrina',
    '3506': '3506 - Colas',
    '3507': '3507 - Enzimas',
    '3601': '3601 - Pólvoras',
    '3602': '3602 - Explosivos',
    '3603': '3603 - Estopins',
    '3604': '3604 - Fogos de artifício',
    '3605': '3605 - Fósforos',
    '3606': '3606 - Ligas pirofóricas',
    '3701': '3701 - Chapas fotográficas',
    '3702': '3702 - Filmes fotográficos',
    '3703': '3703 - Papel fotográfico',
    '3704': '3704 - Chapas preparadas',
    '3705': '3705 - Filmes expostos',
    '3706': '3706 - Filmes cinematográficos',
    '3707': '3707 - Produtos químicos fotográficos',
    '3801': '3801 - Grafite artificial',
    '3802': '3802 - Carvão ativado',
    '3803': '3803 - Tall oil',
    '3804': '3804 - Resíduos de madeira',
    '3805': '3805 - Essências de madeira',
    '3806': '3806 - Colofônia',
    '3807': '3807 - Alcatrão de madeira',
    '3808': '3808 - Inseticidas',
    '3809': '3809 - Acabamentos têxteis',
    '3810': '3810 - Decapantes',
    '3811': '3811 - Aditivos para óleos',
    '3812': '3812 - Aceleradores de vulcanização',
    '3813': '3813 - Preparações para extintores',
    '3814': '3814 - Solventes',
    '3815': '3815 - Iniciadores de reação',
    '3816': '3816 - Cimentos refratários',
    '3817': '3817 - Misturas de alquilbenzenos',
    '3818': '3818 - Elementos químicos dopados',
    '3819': '3819 - Fluidos para freios',
    '3820': '3820 - Anticongelantes',
    '3821': '3821 - Meios de cultura',
    '3822': '3822 - Reagentes de diagnóstico',
    '3823': '3823 - Ácidos graxos',
    '3824': '3824 - Aglutinantes',
    '3825': '3825 - Resíduos químicos',
    '3826': '3826 - Biodiesel',
    '3827': '3827 - Outros produtos químicos',
    
    // Seção VII - Plásticos e borrachas
    '3901': '3901 - Polímeros de etileno',
    '3902': '3902 - Polímeros de propileno',
    '3903': '3903 - Polímeros de estireno',
    '3904': '3904 - Polímeros de cloreto de vinila',
    '3905': '3905 - Polímeros de acetato de vinila',
    '3906': '3906 - Polímeros acrílicos',
    '3907': '3907 - Poliacetais',
    '3908': '3908 - Poliamidas',
    '3909': '3909 - Resinas amínicas',
    '3910': '3910 - Silicones',
    '3911': '3911 - Resinas de petróleo',
    '3912': '3912 - Celulose',
    '3913': '3913 - Polímeros naturais',
    '3914': '3914 - Permutadores de íons',
    '3915': '3915 - Resíduos de plásticos',
    '3916': '3916 - Monofilamentos',
    '3917': '3917 - Tubos de plástico',
    '3918': '3918 - Revestimentos de plástico',
    '3919': '3919 - Folhas auto-adesivas',
    '3920': '3920 - Chapas de plástico',
    '3921': '3921 - Outras chapas de plástico',
    '3922': '3922 - Banheiras',
    '3923': '3923 - Artigos de transporte',
    '3924': '3924 - Artigos de mesa',
    '3925': '3925 - Artigos para construção',
    '3926': '3926 - Outros artigos de plástico',
    '4001': '4001 - Borracha natural',
    '4002': '4002 - Borracha sintética',
    '4003': '4003 - Borracha regenerada',
    '4004': '4004 - Resíduos de borracha',
    '4005': '4005 - Borracha não vulcanizada',
    '4006': '4006 - Borracha vulcanizada',
    '4007': '4007 - Fios de borracha',
    '4008': '4008 - Chapas de borracha',
    '4009': '4009 - Tubos de borracha',
    '4010': '4010 - Correias de borracha',
    '4011': '4011 - Pneumáticos',
    '4012': '4012 - Pneumáticos usados',
    '4013': '4013 - Câmaras de ar',
    '4014': '4014 - Artigos de borracha',
    '4015': '4015 - Artigos de vestuário',
    '4016': '4016 - Outros artigos de borracha',
    '4017': '4017 - Borracha endurecida',
};

// ========== FUNÇÃO PARA OBTER GRUPO NCM PELOS 4 PRIMEIROS DÍGITOS ==========
function getGrupoNCM(codigoNCM) {
    if (!codigoNCM || codigoNCM.length < 4) return 'NCM incompleto (necessário 4 dígitos)';
    
    // Pega os 4 primeiros dígitos
    const quatroPrimeiros = codigoNCM.toString().substring(0, 4);
    
    // Busca no mapeamento
    if (gruposNCM[quatroPrimeiros]) {
        return gruposNCM[quatroPrimeiros];
    }
    
    // Se não encontrou, tenta com os 2 primeiros dígitos como fallback
    const doisPrimeiros = codigoNCM.toString().substring(0, 2);
    const gruposPorDoisDigitos = {
        '01': `01 - Animais vivos (${quatroPrimeiros} - Específico não mapeado)`,
        '02': `02 - Carnes (${quatroPrimeiros} - Específico não mapeado)`,
        '03': `03 - Peixes e crustáceos (${quatroPrimeiros} - Específico não mapeado)`,
        '04': `04 - Laticínios (${quatroPrimeiros} - Específico não mapeado)`,
        '05': `05 - Produtos de origem animal (${quatroPrimeiros} - Específico não mapeado)`,
        '06': `06 - Plantas vivas (${quatroPrimeiros} - Específico não mapeado)`,
        '07': `07 - Legumes e hortaliças (${quatroPrimeiros} - Específico não mapeado)`,
        '08': `08 - Frutas (${quatroPrimeiros} - Específico não mapeado)`,
        '09': `09 - Café, chá e especiarias (${quatroPrimeiros} - Específico não mapeado)`,
        '10': `10 - Cereais (${quatroPrimeiros} - Específico não mapeado)`,
        '11': `11 - Produtos de moagem (${quatroPrimeiros} - Específico não mapeado)`,
        '12': `12 - Sementes e frutos oleaginosos (${quatroPrimeiros} - Específico não mapeado)`,
        '13': `13 - Gomas e resinas (${quatroPrimeiros} - Específico não mapeado)`,
        '14': `14 - Matérias vegetais (${quatroPrimeiros} - Específico não mapeado)`,
        '15': `15 - Gorduras e óleos (${quatroPrimeiros} - Específico não mapeado)`,
        '16': `16 - Preparações de carnes (${quatroPrimeiros} - Específico não mapeado)`,
        '17': `17 - Açúcares (${quatroPrimeiros} - Específico não mapeado)`,
        '18': `18 - Cacau e chocolate (${quatroPrimeiros} - Específico não mapeado)`,
        '19': `19 - Preparações à base de cereais (${quatroPrimeiros} - Específico não mapeado)`,
        '20': `20 - Preparações de legumes e frutas (${quatroPrimeiros} - Específico não mapeado)`,
        '21': `21 - Preparações alimentícias diversas (${quatroPrimeiros} - Específico não mapeado)`,
        '22': `22 - Bebidas (${quatroPrimeiros} - Específico não mapeado)`,
        '23': `23 - Resíduos alimentícios (${quatroPrimeiros} - Específico não mapeado)`,
        '24': `24 - Tabaco (${quatroPrimeiros} - Específico não mapeado)`,
        '25': `25 - Produtos minerais (${quatroPrimeiros} - Específico não mapeado)`,
        '26': `26 - Minérios (${quatroPrimeiros} - Específico não mapeado)`,
        '27': `27 - Combustíveis minerais (${quatroPrimeiros} - Específico não mapeado)`,
        '28': `28 - Produtos químicos inorgânicos (${quatroPrimeiros} - Específico não mapeado)`,
        '29': `29 - Produtos químicos orgânicos (${quatroPrimeiros} - Específico não mapeado)`,
        '30': `30 - Produtos farmacêuticos (${quatroPrimeiros} - Específico não mapeado)`,
        '31': `31 - Fertilizantes (${quatroPrimeiros} - Específico não mapeado)`,
        '32': `32 - Corantes e pigmentos (${quatroPrimeiros} - Específico não mapeado)`,
        '33': `33 - Óleos essenciais (${quatroPrimeiros} - Específico não mapeado)`,
        '34': `34 - Sabões e detergentes (${quatroPrimeiros} - Específico não mapeado)`,
        '35': `35 - Colas e enzimas (${quatroPrimeiros} - Específico não mapeado)`,
        '36': `36 - Explosivos (${quatroPrimeiros} - Específico não mapeado)`,
        '37': `37 - Produtos fotográficos (${quatroPrimeiros} - Específico não mapeado)`,
        '38': `38 - Produtos químicos diversos (${quatroPrimeiros} - Específico não mapeado)`,
        '39': `39 - Plásticos (${quatroPrimeiros} - Específico não mapeado)`,
        '40': `40 - Borrachas (${quatroPrimeiros} - Específico não mapeado)`,
        '41': `41 - Peles e couros (${quatroPrimeiros} - Específico não mapeado)`,
        '42': `42 - Artigos de couro (${quatroPrimeiros} - Específico não mapeado)`,
        '43': `43 - Peles com pelo (${quatroPrimeiros} - Específico não mapeado)`,
        '44': `44 - Madeira (${quatroPrimeiros} - Específico não mapeado)`,
        '45': `45 - Cortiça (${quatroPrimeiros} - Específico não mapeado)`,
        '46': `46 - Obras de palha (${quatroPrimeiros} - Específico não mapeado)`,
        '47': `47 - Pastas de madeira (${quatroPrimeiros} - Específico não mapeado)`,
        '48': `48 - Papel (${quatroPrimeiros} - Específico não mapeado)`,
        '49': `49 - Livros (${quatroPrimeiros} - Específico não mapeado)`,
        '50': `50 - Seda (${quatroPrimeiros} - Específico não mapeado)`,
        '51': `51 - Lã (${quatroPrimeiros} - Específico não mapeado)`,
        '52': `52 - Algodão (${quatroPrimeiros} - Específico não mapeado)`,
        '53': `53 - Fibras têxteis (${quatroPrimeiros} - Específico não mapeado)`,
        '54': `54 - Filamentos sintéticos (${quatroPrimeiros} - Específico não mapeado)`,
        '55': `55 - Fibras sintéticas (${quatroPrimeiros} - Específico não mapeado)`,
        '56': `56 - Pastas têxteis (${quatroPrimeiros} - Específico não mapeado)`,
        '57': `57 - Tapetes (${quatroPrimeiros} - Específico não mapeado)`,
        '58': `58 - Tecidos especiais (${quatroPrimeiros} - Específico não mapeado)`,
        '59': `59 - Tecidos impregnados (${quatroPrimeiros} - Específico não mapeado)`,
        '60': `60 - Tecidos de malha (${quatroPrimeiros} - Específico não mapeado)`,
        '61': `61 - Vestuário de malha (${quatroPrimeiros} - Específico não mapeado)`,
        '62': `62 - Vestuário (${quatroPrimeiros} - Específico não mapeado)`,
        '63': `63 - Outros artefatos têxteis (${quatroPrimeiros} - Específico não mapeado)`,
        '64': `64 - Calçados (${quatroPrimeiros} - Específico não mapeado)`,
        '65': `65 - Chapéus (${quatroPrimeiros} - Específico não mapeado)`,
        '66': `66 - Guarda-chuvas (${quatroPrimeiros} - Específico não mapeado)`,
        '67': `67 - Penas (${quatroPrimeiros} - Específico não mapeado)`,
        '68': `68 - Obras de pedra (${quatroPrimeiros} - Específico não mapeado)`,
        '69': `69 - Cerâmica (${quatroPrimeiros} - Específico não mapeado)`,
        '70': `70 - Vidro (${quatroPrimeiros} - Específico não mapeado)`,
        '71': `71 - Pedras preciosas (${quatroPrimeiros} - Específico não mapeado)`,
        '72': `72 - Ferro e aço (${quatroPrimeiros} - Específico não mapeado)`,
        '73': `73 - Artigos de ferro (${quatroPrimeiros} - Específico não mapeado)`,
        '74': `74 - Cobre (${quatroPrimeiros} - Específico não mapeado)`,
        '75': `75 - Níquel (${quatroPrimeiros} - Específico não mapeado)`,
        '76': `76 - Alumínio (${quatroPrimeiros} - Específico não mapeado)`,
        '78': `78 - Chumbo (${quatroPrimeiros} - Específico não mapeado)`,
        '79': `79 - Zinco (${quatroPrimeiros} - Específico não mapeado)`,
        '80': `80 - Estanho (${quatroPrimeiros} - Específico não mapeado)`,
        '81': `81 - Metais diversos (${quatroPrimeiros} - Específico não mapeado)`,
        '82': `82 - Ferramentas (${quatroPrimeiros} - Específico não mapeado)`,
        '83': `83 - Artigos diversos de metais (${quatroPrimeiros} - Específico não mapeado)`,
        '84': `84 - Máquinas (${quatroPrimeiros} - Específico não mapeado)`,
        '85': `85 - Equipamentos elétricos (${quatroPrimeiros} - Específico não mapeado)`,
        '86': `86 - Veículos ferroviários (${quatroPrimeiros} - Específico não mapeado)`,
        '87': `87 - Veículos automóveis (${quatroPrimeiros} - Específico não mapeado)`,
        '88': `88 - Aeronaves (${quatroPrimeiros} - Específico não mapeado)`,
        '89': `89 - Embarcações (${quatroPrimeiros} - Específico não mapeado)`,
        '90': `90 - Instrumentos de óptica (${quatroPrimeiros} - Específico não mapeado)`,
        '91': `91 - Relógios (${quatroPrimeiros} - Específico não mapeado)`,
        '92': `92 - Instrumentos musicais (${quatroPrimeiros} - Específico não mapeado)`,
        '93': `93 - Armas (${quatroPrimeiros} - Específico não mapeado)`,
        '94': `94 - Móveis (${quatroPrimeiros} - Específico não mapeado)`,
        '95': `95 - Brinquedos (${quatroPrimeiros} - Específico não mapeado)`,
        '96': `96 - Manufaturas diversas (${quatroPrimeiros} - Específico não mapeado)`,
        '97': `97 - Obras de arte (${quatroPrimeiros} - Específico não mapeado)`
    };
    
    return gruposPorDoisDigitos[doisPrimeiros] || `Grupo ${quatroPrimeiros} - NCM não classificado`;
}

// ========== FUNÇÃO PARA FORMATAR GRUPO COM CORES ==========
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}

function formatarGrupoComCor(grupo) {
    // Extrai os primeiros 2 dígitos do grupo para definir a cor
    const doisPrimeiros = grupo.substring(0, 2);
    
    // Mapeamento de cores por seção (baseado nos 2 primeiros dígitos)
    const coresPorSecao = {
        '01': '#FFE5E5', '02': '#FFE5E5', '03': '#FFE5E5', '04': '#FFE5E5', '05': '#FFE5E5', // Animais
        '06': '#E5FFE5', '07': '#E5FFE5', '08': '#E5FFE5', '09': '#E5FFE5', '10': '#E5FFE5', // Vegetais
        '11': '#E5FFE5', '12': '#E5FFE5', '13': '#E5FFE5', '14': '#E5FFE5', // Vegetais
        '15': '#FFF0E5', // Gorduras
        '16': '#FFE5F0', '17': '#FFE5F0', '18': '#FFE5F0', '19': '#FFE5F0', '20': '#FFE5F0', // Alimentos
        '21': '#FFE5F0', '22': '#FFE5F0', '23': '#FFE5F0', // Alimentos
        '24': '#E5F0FF', // Tabaco
        '25': '#E5F0FF', '26': '#E5F0FF', '27': '#E5F0FF', // Minerais
        '28': '#F0E5FF', '29': '#F0E5FF', '30': '#F0E5FF', '31': '#F0E5FF', '32': '#F0E5FF', // Químicos
        '33': '#F0E5FF', '34': '#F0E5FF', '35': '#F0E5FF', '36': '#F0E5FF', '37': '#F0E5FF',
        '38': '#F0E5FF', '39': '#F0E5FF', '40': '#F0E5FF',
        '41': '#F5E5D5', '42': '#F5E5D5', '43': '#F5E5D5', // Couros
        '44': '#E5D5C5', '45': '#E5D5C5', '46': '#E5D5C5', // Madeira
        '47': '#FFE5D5', '48': '#FFE5D5', '49': '#FFE5D5', // Papel
        '50': '#FFE5E5', '51': '#FFE5E5', '52': '#FFE5E5', '53': '#FFE5E5', '54': '#FFE5E5', // Têxteis
        '55': '#FFE5E5', '56': '#FFE5E5', '57': '#FFE5E5', '58': '#FFE5E5', '59': '#FFE5E5',
        '60': '#FFE5E5', '61': '#FFE5E5', '62': '#FFE5E5', '63': '#FFE5E5',
        '64': '#E5E5FF', '65': '#E5E5FF', '66': '#E5E5FF', '67': '#E5E5FF', // Calçados
        '68': '#E5F5E5', '69': '#E5F5E5', '70': '#E5F5E5', // Pedra e cerâmica
        '71': '#FFF0E5', // Pedras preciosas
        '72': '#E5E5E5', '73': '#E5E5E5', '74': '#E5E5E5', '75': '#E5E5E5', '76': '#E5E5E5', // Metais
        '78': '#E5E5E5', '79': '#E5E5E5', '80': '#E5E5E5', '81': '#E5E5E5', '82': '#E5E5E5',
        '83': '#E5E5E5',
        '84': '#FFE5E5', '85': '#FFE5E5', // Máquinas
        '86': '#E5FFE5', '87': '#E5FFE5', '88': '#E5FFE5', '89': '#E5FFE5', // Transporte
        '90': '#FFE5F0', '91': '#FFE5F0', '92': '#FFE5F0', // Instrumentos
        '93': '#FFE5E5', // Armas
        '94': '#F0F0F0', // Móveis
        '95': '#F5E5F0', // Brinquedos
        '96': '#F0F0F0', // Diversos
        '97': '#F5F0E5' // Arte
    };
    
    let cor = coresPorSecao[doisPrimeiros] || '#F5F5F5';
    let corEscura = false;
    
    // Determina se a cor é escura
    const rgb = hexToRgb(cor);
    if (rgb) {
        const luminosidade = (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b);
        corEscura = luminosidade < 128;
    }
    
    // Formata o texto para exibição mais amigável
    let textoExibicao = grupo;
    
    // Se for um grupo genérico de 2 dígitos, mostra de forma mais limpa
    if (textoExibicao.includes('Específico não mapeado')) {
        textoExibicao = textoExibicao.split(' - ')[0];
    }
    
    return {
        cor: cor,
        textoCor: corEscura ? 'white' : '#333',
        html: `<span style="background-color: ${cor}; color: ${corEscura ? 'white' : '#333'}; padding: 6px 10px; border-radius: 6px; font-size: 12px; display: inline-block; font-weight: 500; max-width: 300px; white-space: normal;">${textoExibicao}</span>`
    };
}

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

    const cidadeNome = dados.cidade?.nome || 'Não informado';
    const uf = dados.uf || 'Não informado';
    const porteDescricao = dados.porte?.descricao || 'Não informado';
    
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
    
    let email = dados.email || 'Não informado';
    if (dados.email === null || dados.email === undefined || dados.email === '') {
        email = 'Não informado';
    }

    const janelaImpressao = window.open('', '_blank');
    
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
    } else if (nomeAba === 'senha') {
        document.querySelectorAll('.tab-button')[4].classList.add('active');
        document.getElementById('abaSenha').classList.add('active');
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

    // Usar Map para garantir NCMs únicos
    const resultadosMap = new Map();
    
    baseNCM.forEach(item => {
        if (ncmTem8Digitos(item.codigo)) {
            const codigoMatch = item.codigo.includes(termo);
            const descMatch = item.descricao.toLowerCase().includes(termo);
            
            if (codigoMatch || descMatch) {
                // Garantir que não adiciona códigos duplicados
                if (!resultadosMap.has(item.codigo)) {
                    resultadosMap.set(item.codigo, item);
                }
            }
        }
    });

    // Converter Map para array e limitar a 50 resultados
    const resultados = Array.from(resultadosMap.values()).slice(0, 50);

    statusNcm.innerText = `${resultados.length} resultados únicos para "${termo}"`;

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
    window.ultimoCnpjConsultado = dados;
    
    const dataAbertura = dados.data_inicio_atividade ? new Date(dados.data_inicio_atividade).toLocaleDateString('pt-BR') : 'Não informado';
    const dataSituacao = dados.data_situacao_cadastral ? new Date(dados.data_situacao_cadastral).toLocaleDateString('pt-BR') : 'Não informado';
    
    let telefone1 = 'Não informado';
    let telefone2 = 'Não informado';
    
    if (dados.ddd_telefone_1) {
        telefone1 = formatarCampoCNPJ(dados.ddd_telefone_1, 'telefone');
    }
    
    if (dados.ddd_telefone_2) {
        telefone2 = formatarCampoCNPJ(dados.ddd_telefone_2, 'telefone');
    }
    
    const email = dados.email || 'Não informado';
    
    const regimeTributario = dados.regime_tributario || [];
    const ultimoRegime = regimeTributario.length > 0 ? regimeTributario[regimeTributario.length - 1] : null;
    
    return `
        <div class="cnpj-info">
            <div class="cnpj-section">
                <h4>📋 DADOS PRINCIPAIS</h4>
                <p><strong>CNPJ:</strong> ${dados.cnpj || 'Não informado'}</p>
                <p><strong>Razão Social:</strong> ${dados.razao_social || 'Não informado'}</p>
                <p><strong>Nome Fantasia:</strong> ${dados.nome_fantasia || 'Não informado'}</p>
                <p><strong>Natureza Jurídica:</strong> ${dados.natureza_juridica || 'Não informado'} (${dados.codigo_natureza_juridica || ''})</p>
                <p><strong>Porte:</strong> ${dados.porte || 'Não informado'} (Código: ${dados.codigo_porte || ''})</p>
                <p><strong>Capital Social:</strong> ${dados.capital_social ? formatarCampoCNPJ(dados.capital_social, 'capitalSocial') : 'Não informado'}</p>
            </div>
            
            <div class="cnpj-section">
                <h4>⚖️ SITUAÇÃO CADASTRAL</h4>
                <p><strong>Situação:</strong> ${dados.descricao_situacao_cadastral || 'Não informado'} (Código: ${dados.situacao_cadastral || ''})</p>
                <p><strong>Data da Situação:</strong> ${dataSituacao}</p>
                <p><strong>Motivo:</strong> ${dados.descricao_motivo_situacao_cadastral || 'Não informado'} (Código: ${dados.motivo_situacao_cadastral || ''})</p>
                <p><strong>Situação Especial:</strong> ${dados.situacao_especial || 'Não informado'}</p>
                <p><strong>Data Situação Especial:</strong> ${dados.data_situacao_especial ? new Date(dados.data_situacao_especial).toLocaleDateString('pt-BR') : 'Não informado'}</p>
            </div>
            
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
            
            <div class="cnpj-section">
                <h4>📞 CONTATO</h4>
                <p><strong>Telefone 1:</strong> ${telefone1}</p>
                <p><strong>Telefone 2:</strong> ${telefone2}</p>
                <p><strong>Fax:</strong> ${dados.ddd_fax ? formatarCampoCNPJ(dados.ddd_fax, 'telefone') : 'Não informado'}</p>
                <p><strong>Email:</strong> ${email}</p>
            </div>
            
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

// ========== TRATAMENTO DE ENCODING E CARACTERES ESPECIAIS ==========

/**
 * Detecta e corrige problemas de encoding em textos
 * @param {string} texto - Texto com possíveis problemas de encoding
 * @returns {string} - Texto corrigido
 */
function corrigirEncoding(texto) {
    if (!texto || typeof texto !== 'string') return texto;
    
    // Mapeamento de caracteres comuns corrompidos (Windows-1252 -> UTF-8)
    const mapaCorrecao = {
        // Acentuação comum corrompida
        'Ã¡': 'á', 'Ã¢': 'â', 'Ã£': 'ã', 'Ã¤': 'ä', 'Ã ': 'à',
        'Ã©': 'é', 'Ãª': 'ê', 'Ã«': 'ë',
        'Ã­': 'í', 'Ã®': 'î', 'Ã¯': 'ï',
        'Ã³': 'ó', 'Ã´': 'ô', 'Ãµ': 'õ', 'Ã¶': 'ö',
        'Ãº': 'ú', 'Ã»': 'û', 'Ã¼': 'ü',
        'Ã§': 'ç', 'Ãƒ': 'ã', 'Ã‰': 'é', 'Ã§Ã£o': 'ção',
        
        // Caracteres especiais corrompidos
        'Â°': '°', 'Âº': 'º', 'Âª': 'ª', 'Â£': '£',
        'â‚¬': '€', 'â„¢': '™', 'Â®': '®', 'Â©': '©',
        'ÃŸ': 'ß', 'Ã†': 'Æ', 'Ã˜': 'Ø',
        
        // Espaços e pontuação
        'Ã‚': 'Â', 'ÃƒÂ': 'ã', 'Ã¢â‚¬â€œ': '—',
        'Ã¢â‚¬â„¢': "'", 'Ã¢â‚¬Ëœ': "'", 'Ã¢â‚¬Å“': '"',
        'Ã¢â‚¬Â': '"', 'Ã¢â€šÂ¬': '€',
        
        // Casos específicos do seu arquivo
        'algod?o': 'algodão',
        'cora?': 'coração',
        '?': 'ã', // Corrige interrogação mal formatada como "ã" em alguns casos
    };
    
    // Aplica correções específicas
    let textoCorrigido = texto;
    for (const [corrompido, corrigido] of Object.entries(mapaCorrecao)) {
        textoCorrigido = textoCorrigido.replace(new RegExp(corrompido, 'g'), corrigido);
    }
    
    // Tenta corrigir padrões comuns de double-encoding
    try {
        // Se parece ser UTF-8 duplicado, tenta decodificar
        if (textoCorrigido.includes('Ã')) {
            // Converte para bytes e tenta redecodificar
            const bytes = [];
            for (let i = 0; i < textoCorrigido.length; i++) {
                const code = textoCorrigido.charCodeAt(i);
                if (code > 127 && code < 256) {
                    bytes.push(code);
                }
            }
            
            if (bytes.length > 0) {
                const decoder = new TextDecoder('utf-8');
                const decoded = decoder.decode(new Uint8Array(bytes));
                if (decoded && !decoded.includes('�')) {
                    textoCorrigido = decoded;
                }
            }
        }
    } catch (e) {
        console.warn('Erro ao tentar corrigir double-encoding:', e);
    }
    
    return textoCorrigido;
}

/**
 * Normaliza caracteres especiais (opcional - para buscas mais flexíveis)
 * @param {string} texto - Texto para normalizar
 * @returns {string} - Texto sem acentos e em minúsculo
 */
function normalizarTexto(texto) {
    if (!texto || typeof texto !== 'string') return texto;
    
    // Primeiro corrige encoding se necessário
    let normalizado = corrigirEncoding(texto);
    
    // Remove acentos
    normalizado = normalizado.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    
    // Converte para minúsculo
    normalizado = normalizado.toLowerCase();
    
    return normalizado;
}

/**
 * Limpa e valida string, removendo caracteres inválidos
 * @param {string} texto - Texto para limpar
 * @returns {string} - Texto limpo
 */
function limparString(texto) {
    if (!texto || typeof texto !== 'string') return '';
    
    // Remove caracteres de controle
    let limpo = texto.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
    
    // Remove BOM (Byte Order Mark)
    limpo = limpo.replace(/^\uFEFF/, '');
    
    // Remove caracteres nulos
    limpo = limpo.replace(/\0/g, '');
    
    // Corrige encoding
    limpo = corrigirEncoding(limpo);
    
    return limpo.trim();
}

/**
 * Detecta o encoding de um texto/base64
 * @param {string|ArrayBuffer} dados - Dados para detectar encoding
 * @returns {string} - Encoding detectado
 */
function detectarEncoding(dados) {
    if (!dados) return 'UTF-8';
    
    // Converte para string se necessário
    let texto = typeof dados === 'string' ? dados : new TextDecoder().decode(dados);
    
    // Verifica padrões comuns de encoding errado
    if (texto.includes('Ã') && (texto.includes('§') || texto.includes('£'))) {
        return 'WINDOWS-1252';
    }
    
    if (texto.includes('Ã') && texto.includes('©')) {
        return 'ISO-8859-1';
    }
    
    if (texto.includes('�')) {
        return 'BROKEN';
    }
    
    return 'UTF-8';
}

/**
 * Lê CSV com suporte a múltiplos encodings
 * @param {File} file - Arquivo CSV
 * @returns {Promise<Array>} - Array de produtos
 */
async function lerCSVComEncoding(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (event) => {
            let content = event.target.result;
            let produtos = [];
            const produtosSet = new Set();
            const duplicatas = [];
            
            // Detecta e tenta corrigir encoding
            const encoding = detectarEncoding(content);
            
            if (encoding === 'BROKEN' || content.includes('�')) {
                // Tenta ler o arquivo original em diferentes encodings
                const originalBytes = event.target.result;
                const byteArray = new Uint8Array(originalBytes);
                
                // Tenta diferentes decodificadores
                const encodings = ['utf-8', 'windows-1252', 'iso-8859-1'];
                let melhorConteudo = null;
                let menosCaracteresRuins = Infinity;
                
                for (const enc of encodings) {
                    try {
                        const decoder = new TextDecoder(enc);
                        const decodificado = decoder.decode(byteArray);
                        const charsRuins = (decodificado.match(/[�\x00-\x08\x0B\x0C\x0E-\x1F]/g) || []).length;
                        
                        if (charsRuins < menosCaracteresRuins) {
                            menosCaracteresRuins = charsRuins;
                            melhorConteudo = decodificado;
                        }
                    } catch (e) {
                        continue;
                    }
                }
                
                if (melhorConteudo) {
                    content = melhorConteudo;
                }
            }
            
            // Divide em linhas
            const lines = content.split(/\r?\n/);
            
            for (let i = 0; i < lines.length; i++) {
                let line = lines[i].trim();
                if (!line) continue;
                
                // Tenta diferentes separadores
                let fields;
                if (line.includes(';')) {
                    fields = line.split(';');
                } else if (line.includes(',')) {
                    fields = line.split(',');
                } else if (line.includes('\t')) {
                    fields = line.split('\t');
                } else {
                    fields = [line];
                }
                
                let nomeProduto = fields[0];
                
                // Limpa e corrige o nome do produto
                nomeProduto = limparString(nomeProduto);
                
                // Remove aspas se existirem
                nomeProduto = nomeProduto.replace(/^["']|["']$/g, '');
                
                // Corrige problemas específicos de encoding
                nomeProduto = corrigirEncoding(nomeProduto);
                
                if (nomeProduto && nomeProduto !== '') {
                    const nomeLower = normalizarTexto(nomeProduto);
                    
                    if (!produtosSet.has(nomeLower)) {
                        produtosSet.add(nomeLower);
                        produtos.push({ 
                            nome: nomeProduto,
                            nomeNormalizado: nomeLower,
                            linha: i + 1 
                        });
                    } else {
                        duplicatas.push(`Linha ${i + 1}: ${nomeProduto}`);
                    }
                }
            }
            
            resolve({ produtos, duplicatas: duplicatas.length });
        };
        
        reader.onerror = reject;
        
        // Tenta ler como texto primeiro
        try {
            reader.readAsText(file, 'UTF-8');
        } catch (e) {
            // Se falhar, tenta como ArrayBuffer para detectar encoding
            reader.readAsArrayBuffer(file);
        }
    });
}

// ========== FUNÇÃO DE BUSCA MELHORADA ==========
function buscarNCMProdutoMelhorado(nomeProduto) {
    if (!baseNCM || baseNCM.length === 0) {
        return { encontrado: false, ncm: null, descricao: null };
    }
    
    // Corrige encoding do produto
    const produtoCorrigido = corrigirEncoding(nomeProduto);
    const termo = produtoCorrigido.toLowerCase();
    const termoNormalizado = normalizarTexto(produtoCorrigido);
    
    // Palavras-chave do produto (remove palavras comuns)
    const palavrasChave = termoNormalizado
        .split(/\s+/)
        .filter(p => p.length > 2 && !['com', 'para', 'dos', 'das', 'de', 'da', 'do', 'e', 'a', 'o'].includes(p));
    
    // Usar Map para garantir resultados únicos
    const resultadosMap = new Map();
    
    // Busca exata (com e sem normalização)
    baseNCM.forEach(item => {
        if (ncmTem8Digitos(item.codigo)) {
            const descNormalizada = normalizarTexto(item.descricao);
            const descOriginal = item.descricao.toLowerCase();
            
            // Busca exata normalizada
            if (descNormalizada.includes(termoNormalizado)) {
                if (!resultadosMap.has(item.codigo)) {
                    resultadosMap.set(item.codigo, item);
                }
            }
            // Busca exata original (com encoding corrigido)
            else if (descOriginal.includes(termo)) {
                if (!resultadosMap.has(item.codigo)) {
                    resultadosMap.set(item.codigo, item);
                }
            }
        }
    });
    
    // Se não encontrar, busca por palavras-chave
    if (resultadosMap.size === 0 && palavrasChave.length > 0) {
        baseNCM.forEach(item => {
            if (ncmTem8Digitos(item.codigo)) {
                const descNormalizada = normalizarTexto(item.descricao);
                const matchCount = palavrasChave.filter(palavra => descNormalizada.includes(palavra)).length;
                
                // Se pelo menos 50% das palavras-chave encontradas
                if (matchCount >= palavrasChave.length * 0.5) {
                    if (!resultadosMap.has(item.codigo)) {
                        resultadosMap.set(item.codigo, item);
                    }
                }
            }
        });
    }
    
    // Pega o melhor resultado (primeiro)
    const melhorResultado = Array.from(resultadosMap.values())[0];
    
    if (melhorResultado) {
        return {
            encontrado: true,
            ncm: melhorResultado.codigo,
            descricao: melhorResultado.descricao,
            nomeOriginal: nomeProduto,
            nomeCorrigido: produtoCorrigido
        };
    }
    
    return { 
        encontrado: false, 
        ncm: null, 
        descricao: null,
        nomeOriginal: nomeProduto,
        nomeCorrigido: produtoCorrigido
    };
}

// ========== SOBRESCREVER A FUNÇÃO DE LEITURA CSV ORIGINAL ==========
// Substitui a função lerCSV original pela versão com encoding
const lerCSVOriginal = window.lerCSV;
window.lerCSV = lerCSVComEncoding;

// ========== FUNÇÃO PARA CORRIGIR RESULTADOS EXISTENTES ==========
function corrigirResultadosExistentes() {
    if (!window.ultimosResultados || window.ultimosResultados.length === 0) {
        alert('Nenhum resultado para corrigir!');
        return;
    }
    
    const resultadosCorrigidos = window.ultimosResultados.map(r => ({
        ...r,
        nome: corrigirEncoding(r.nome),
        descricao: r.descricao ? corrigirEncoding(r.descricao) : r.descricao
    }));
    
    exibirTabelaResultados(resultadosCorrigidos);
    window.ultimosResultados = resultadosCorrigidos;
    
    statusImportacao.innerText = "✅ Resultados corrigidos!";
    statusImportacao.style.color = "green";
}

// ========== ADICIONAR BOTÃO DE CORREÇÃO ==========
// Adicionar botão na interface de importação
function adicionarBotaoCorrecao() {
    const container = document.querySelector('#abaImportar .import-section');
    if (container && !document.getElementById('btnCorrigirResultados')) {
        const btnCorrigir = document.createElement('button');
        btnCorrigir.id = 'btnCorrigirResultados';
        btnCorrigir.innerHTML = '🔧 Corrigir Caracteres dos Resultados';
        btnCorrigir.style.marginLeft = '10px';
        btnCorrigir.style.background = '#f59e0b';
        btnCorrigir.onclick = corrigirResultadosExistentes;
        
        const btnExportar = document.getElementById('btnExportarCSV');
        if (btnExportar) {
            btnExportar.parentNode.insertBefore(btnCorrigir, btnExportar.nextSibling);
        }
    }
}

// Executar após carregar a página
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', adicionarBotaoCorrecao);
} else {
    adicionarBotaoCorrecao();
}

// ========== FUNÇÕES DE IMPORTAÇÃO CSV ==========
function lerCSV(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (event) => {
            const lines = event.target.result.split('\n');
            const produtos = [];
            const produtosSet = new Set(); // Para controlar produtos únicos
            const duplicatas = [];
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line) {
                    const fields = line.split(';');
                    let nomeProduto = fields[0].trim();
                    
                    // Remove aspas se existirem
                    nomeProduto = nomeProduto.replace(/^["']|["']$/g, '');
                    
                    if (nomeProduto) {
                        const nomeLower = nomeProduto.toLowerCase();
                        
                        // Verifica se já existe
                        if (!produtosSet.has(nomeLower)) {
                            produtosSet.add(nomeLower);
                            produtos.push({ 
                                nome: nomeProduto,
                                linha: i + 1 
                            });
                        } else {
                            duplicatas.push(`Linha ${i + 1}: ${nomeProduto}`);
                        }
                    }
                }
            }
            
            // Mostra aviso se houver duplicatas
            if (duplicatas.length > 0) {
                console.log('⚠️ Produtos duplicados ignorados:', duplicatas);
            }
            
            resolve({ produtos, duplicatas: duplicatas.length });
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
    
    // Usar Map para garantir resultados únicos por código NCM
    const resultadosMap = new Map();
    
    // Primeiro tenta busca exata
    baseNCM.forEach(item => {
        if (ncmTem8Digitos(item.codigo) && item.descricao.toLowerCase().includes(termo)) {
            if (!resultadosMap.has(item.codigo)) {
                resultadosMap.set(item.codigo, item);
            }
        }
    });
    
    // Se não encontrar, tenta com palavras individuais
    if (resultadosMap.size === 0 && palavras.length > 0) {
        baseNCM.forEach(item => {
            if (ncmTem8Digitos(item.codigo)) {
                const descLower = item.descricao.toLowerCase();
                if (palavras.some(palavra => descLower.includes(palavra))) {
                    if (!resultadosMap.has(item.codigo)) {
                        resultadosMap.set(item.codigo, item);
                    }
                }
            }
        });
    }
    
    // Pega o primeiro resultado único
    const primeiroResultado = Array.from(resultadosMap.values())[0];
    
    if (primeiroResultado) {
        return {
            encontrado: true,
            ncm: primeiroResultado.codigo,
            descricao: primeiroResultado.descricao
        };
    }
    
    return { encontrado: false, ncm: null, descricao: null };
}

// ========== FUNÇÃO EXIBIR TABELA RESULTADOS COM GRUPOS NCM ==========
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
            <table style="width:100%; border-collapse: collapse; min-width: 900px;">
                <thead>
                    <tr style="background: #2563eb; color: white;">
                        <th style="padding: 12px; text-align: left;">Produto</th>
                        <th style="padding: 12px; text-align: left;">NCM (8 dígitos)</th>
                        <th style="padding: 12px; text-align: left;">Grupo NCM</th>
                        <th style="padding: 12px; text-align: left;">Descrição NCM</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    resultados.forEach(r => {
        const corLinha = r.encontrado ? '' : 'background-color: #fff0f0;';
        let grupoHtml = '';
        
        if (r.encontrado && r.ncm) {
            const grupoTexto = getGrupoNCM(r.ncm);
            const grupoFormatado = formatarGrupoComCor(grupoTexto);
            grupoHtml = grupoFormatado.html;
        } else {
            grupoHtml = '<span style="color: #999; font-style: italic;">NCM não encontrado</span>';
        }
        
        html += `
            <tr style="border-bottom: 1px solid #e5e7eb; ${corLinha}">
                <td style="padding: 10px;">${r.nome}</td>
                <td style="padding: 10px; font-weight: bold; font-family: monospace; ${r.encontrado ? 'color: #2563eb;' : 'color: #999;'}">
                    ${r.ncm || '--------'}
                </td>
                <td style="padding: 10px;">
                    ${grupoHtml}
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
        const { produtos, duplicatas } = await lerCSV(file);
        
        if (produtos.length === 0) {
            statusImportacao.innerText = "❌ Nenhum produto encontrado no arquivo!";
            statusImportacao.style.color = "red";
            return;
        }
        
        let mensagem = `🔍 Buscando NCMs para ${produtos.length} produtos únicos...`;
        if (duplicatas > 0) {
            mensagem = `🔍 Buscando NCMs para ${produtos.length} produtos únicos (${duplicatas} duplicatas ignoradas)...`;
        }
        statusImportacao.innerText = mensagem;
        
        const resultados = produtos.map(produto => ({
            ...produto,
            ...buscarNCMProduto(produto.nome)
        }));
        
        exibirTabelaResultados(resultados);
        
        const encontrados = resultados.filter(r => r.encontrado).length;
        let msgFinal = `✅ Processamento concluído! ${encontrados} de ${resultados.length} produtos encontrados.`;
        if (duplicatas > 0) {
            msgFinal += ` (${duplicatas} duplicatas ignoradas)`;
        }
        statusImportacao.innerText = msgFinal;
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
    const cabecalho = ['Produto', 'NCM (8 dígitos)', 'Grupo NCM', 'Descrição NCM', 'Status'];
    
    const linhas = resultados.map(r => {
        const grupoNCM = r.encontrado && r.ncm ? getGrupoNCM(r.ncm) : 'NÃO ENCONTRADO';
        return [
            `"${r.nome.replace(/"/g, '""')}"`,
            r.ncm || 'NÃO ENCONTRADO',
            `"${grupoNCM.replace(/"/g, '""')}"`,
            `"${(r.descricao || 'Não encontrado').replace(/"/g, '""')}"`,
            r.encontrado ? 'ENCONTRADO' : 'NÃO ENCONTRADO'
        ];
    });
    
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
// Função para extrair CEST do produto
function extrairCEST(det, ns) {
    // Tenta encontrar CEST diretamente no produto
    let cestTag = det.getElementsByTagNameNS(ns.nfe, "CEST")[0];
    if (!cestTag) {
        cestTag = det.getElementsByTagName("CEST")[0];
    }
    
    if (cestTag) {
        return cestTag.textContent;
    }
    
    // Se não encontrou, tenta dentro do ICMS (pode estar em algumas estruturas)
    let icms = det.getElementsByTagNameNS(ns.nfe, "ICMS")[0];
    if (!icms) {
        icms = det.getElementsByTagName("ICMS")[0];
    }
    
    if (icms) {
        // ICMS pode estar em diferentes tags (ICMS00, ICMS10, ICMS20, etc)
        const icmsFilhos = icms.children;
        for (let i = 0; i < icmsFilhos.length; i++) {
            const filho = icmsFilhos[i];
            // Procura por tag que contém CEST
            let cestTag = filho.getElementsByTagNameNS(ns.nfe, "CEST")[0];
            if (!cestTag) {
                cestTag = filho.getElementsByTagName("CEST")[0];
            }
            
            if (cestTag) {
                return cestTag.textContent;
            }
        }
    }
    
    return 'CEST não informado';
}

// Função para extrair produtos do XML com verificação de duplicatas
function extrairProdutosXML(xmlString, nomeArquivo) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "text/xml");
    
    // Verifica se há erro de parsing
    const parseError = xmlDoc.getElementsByTagName("parsererror");
    if (parseError.length > 0) {
        throw new Error("Erro ao fazer parse do XML");
    }
    
    // Namespace da NFe
    const ns = {
        nfe: "http://www.portalfiscal.inf.br/nfe"
    };
    
    // Função para obter valor de tag com namespace
    function getTagValue(parent, tagName) {
        // Tenta com namespace
        let element = parent.getElementsByTagNameNS(ns.nfe, tagName)[0];
        
        // Se não encontrou, tenta sem namespace
        if (!element) {
            element = parent.getElementsByTagName(tagName)[0];
        }
        
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
    
    // Se não encontrou com namespace, tenta sem namespace
    const produtos = dets.length > 0 ? dets : infNFe.getElementsByTagName("det");
    
    if (produtos.length === 0) {
        throw new Error("Nenhum produto encontrado no XML");
    }
    
    // Usar Map para garantir produtos únicos (por código + nome + NCM)
    const produtosMap = new Map();
    const duplicatas = [];
    
    for (let i = 0; i < produtos.length; i++) {
        // Tenta encontrar prod com namespace
        let prod = produtos[i].getElementsByTagNameNS(ns.nfe, "prod")[0];
        
        // Se não encontrou, tenta sem namespace
        if (!prod) {
            prod = produtos[i].getElementsByTagName("prod")[0];
        }
        
        if (prod) {
            // Extrai o CEST para este produto
            const cest = extrairCEST(produtos[i], ns);
            
            const produto = {
                codigo: getTagValue(prod, "cProd"),
                nome: getTagValue(prod, "xProd"),
                ncm: getTagValue(prod, "NCM"),
                cest: cest,
                cfop: getTagValue(prod, "CFOP"),
                unidade: getTagValue(prod, "uCom"),
                quantidade: getTagValue(prod, "qCom"),
                valorUnitario: getTagValue(prod, "vUnCom"),
                valorTotal: getTagValue(prod, "vProd"),
                ean: getTagValue(prod, "cEAN"),
                eanTrib: getTagValue(prod, "cEANTrib"),
                arquivoOrigem: nomeArquivo
            };
            
            // Criar chave única (código + nome + NCM)
            const chaveUnica = `${produto.codigo}|${produto.nome}|${produto.ncm}`;
            
            // Verificar se produto já existe no Map
            if (!produtosMap.has(chaveUnica)) {
                produtosMap.set(chaveUnica, produto);
            } else {
                duplicatas.push(`Produto duplicado ignorado: ${produto.codigo} - ${produto.nome}`);
            }
        }
    }
    
    // Log de duplicatas encontradas
    if (duplicatas.length > 0) {
        console.log(`⚠️ Encontradas ${duplicatas.length} duplicatas no arquivo ${nomeArquivo}:`, duplicatas);
    }
    
    return Array.from(produtosMap.values());
}

// ========== FUNÇÃO EXIBIR TABELA PRODUTOS XML COM GRUPOS NCM ==========
function exibirTabelaProdutosXml(produtos) {
    window.ultimosProdutosXml = produtos;
    
    let html = `
        <div style="margin-bottom: 15px; padding: 10px; background: #f0f4ff; border-radius: 8px;">
            <strong>Total de produtos únicos:</strong> ${produtos.length}
        </div>
        <div style="overflow-x: auto; max-height: 500px; overflow-y: auto;">
            <table style="width:100%; border-collapse: collapse; min-width: 1800px;">
                <thead style="position: sticky; top: 0; background: #2563eb; color: white;">
                    <tr>
                        <th style="padding: 12px; text-align: left;">Arquivo</th>
                        <th style="padding: 12px; text-align: left;">Código</th>
                        <th style="padding: 12px; text-align: left;">Produto</th>
                        <th style="padding: 12px; text-align: left;">NCM</th>
                        <th style="padding: 12px; text-align: left;">Grupo NCM</th>
                        <th style="padding: 12px; text-align: left;">CEST</th>
                        <th style="padding: 12px; text-align: left;">CFOP</th>
                        <th style="padding: 12px; text-align: left;">GTIN</th>
                        <th style="padding: 12px; text-align: right;">Qtd</th>
                        <th style="padding: 12px; text-align: left;">UN</th>
                        <th style="padding: 12px; text-align: right;">Valor Unit.</th>
                        <th style="padding: 12px; text-align: right;">Total</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    produtos.forEach(p => {
        const quantidade = p.quantidade ? parseFloat(p.quantidade).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0,00';
        const valorUnit = p.valorUnitario ? parseFloat(p.valorUnitario).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0,00';
        const valorTotal = p.valorTotal ? parseFloat(p.valorTotal).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0,00';
        
        // Define estilo para CEST encontrado ou não
        const cestStyle = p.cest && p.cest !== 'CEST não informado' ? 'color: #2563eb; font-weight: bold;' : 'color: #999; font-style: italic;';
        
        // Obter grupo NCM
        let grupoHtml = '';
        if (p.ncm) {
            const grupoTexto = getGrupoNCM(p.ncm);
            const grupoFormatado = formatarGrupoComCor(grupoTexto);
            grupoHtml = grupoFormatado.html;
        } else {
            grupoHtml = '<span style="color: #999; font-style: italic;">NCM não informado</span>';
        }
        
        html += `
            <tr style="border-bottom: 1px solid #e5e7eb;">
                <td style="padding: 10px; font-size: 12px;">${p.arquivoOrigem || '---'}</td>
                <td style="padding: 10px;">${p.codigo || '---'}</td>
                <td style="padding: 10px;">${p.nome || '---'}</td>
                <td style="padding: 10px; font-weight: bold; color: #2563eb;">${p.ncm || '---'}</td>
                <td style="padding: 10px;">${grupoHtml}</td>
                <td style="padding: 10px; ${cestStyle}">${p.cest || '---'}</td>
                <td style="padding: 10px;">${p.cfop || '---'}</td>
                <td style="padding: 10px; font-family: monospace;">${p.ean || '---'}</td>
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
    let totalDuplicatas = 0;
    
    try {
        for (let i = 0; i < xmlFileInput.files.length; i++) {
            const file = xmlFileInput.files[i];
            try {
                const text = await file.text();
                const produtos = extrairProdutosXML(text, file.name);
                
                // Verificar duplicatas entre diferentes arquivos
                for (const produto of produtos) {
                    const chaveUnica = `${produto.codigo}|${produto.nome}|${produto.ncm}`;
                    const existe = todosProdutos.some(p => 
                        `${p.codigo}|${p.nome}|${p.ncm}` === chaveUnica
                    );
                    
                    if (!existe) {
                        todosProdutos.push(produto);
                    } else {
                        totalDuplicatas++;
                        console.log(`⚠️ Produto duplicado entre arquivos ignorado: ${produto.codigo} - ${produto.nome}`);
                    }
                }
                
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
        
        let mensagem = `✅ Processamento concluído! ${todosProdutos.length} produtos únicos encontrados.`;
        if (totalDuplicatas > 0) {
            mensagem += ` (${totalDuplicatas} duplicatas ignoradas)`;
        }
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
        'Grupo NCM',
        'CEST',
        'CFOP',
        'GTIN',
        'Quantidade',
        'Unidade',
        'Valor Unitário',
        'Valor Total'
    ];
    
    const linhas = produtos.map(p => {
        const grupoNCM = p.ncm ? getGrupoNCM(p.ncm) : 'NCM não informado';
        return [
            `"${p.arquivoOrigem || 'NFe'}"`,
            `"${p.codigo || ''}"`,
            `"${(p.nome || '').replace(/"/g, '""')}"`,
            p.ncm || '',
            `"${grupoNCM.replace(/"/g, '""')}"`,
            p.cest || '',
            p.cfop || '',
            p.ean || '',
            p.quantidade ? parseFloat(p.quantidade).toFixed(2).replace('.', ',') : '0,00',
            p.unidade || '',
            p.valorUnitario ? parseFloat(p.valorUnitario).toFixed(2).replace('.', ',') : '0,00',
            p.valorTotal ? parseFloat(p.valorTotal).toFixed(2).replace('.', ',') : '0,00'
        ];
    });
    
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

function buscarSenha() {
    const data = new Date();
    const senha = data.getFullYear() - data.getDate() - data.getHours();
    
    const inputSenha = document.getElementById('senhaInput');
    inputSenha.value = senha;
    
    // Efeito de destaque temporário
    inputSenha.style.transform = 'scale(1.1)';
    setTimeout(() => {
        inputSenha.style.transform = 'scale(1)';
    }, 200);
}

document.getElementById('btnBuscarSenha').addEventListener('click', buscarSenha);
