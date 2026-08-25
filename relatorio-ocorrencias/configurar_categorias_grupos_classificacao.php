<?php
// Configuração estrutural de TQUIM > Ocorrências: categorias hierárquicas,
// grupos por unidade, bloco de Classificação da Qualidade e data de abertura
// automática — feito uma vez em 25/08/2026, registrado aqui pra ficar
// reproduzível/documentado (não é necessário rodar de novo).
//
// PRÉ-REQUISITO MANUAL: o bloco "Classificação da Qualidade" (Plugin Fields)
// precisa ser criado pela tela do GLPI antes de rodar a parte 3 deste script
// — criar PluginFieldsContainer pela API corrompe o campo "itemtypes"
// internamente (bug conhecido, não tem workaround por API). Vá em
// Configurar > Plugins > Fields > Blocos e crie:
//   nome: classificacao | rótulo: Classificação | tipo: Aba (tab)
//   tipo de item: Ticket | entidade: TQUIM > Ocorrências | recursivo: sim
// Anote o ID gerado e ajuste a constante CONTAINER_CLASSIFICACAO abaixo.
//
// Uso: php configurar_categorias_grupos_classificacao.php

require_once '/var/www/html/glpi/src/Glpi/Application/ResourcesChecker.php';
require_once '/var/www/html/glpi/vendor/autoload.php';

use Glpi\Kernel\Kernel;

$kernel = new Kernel();
$kernel->boot();

const ENTIDADE_OCORRENCIAS = 6;
const TEMPLATE_OCORRENCIAS = 4;
const CONTAINER_CLASSIFICACAO = 6; // ajuste para o ID criado manualmente

global $DB;

// 1. Categorias hierárquicas: agrupa por categoria-pai as que têm nome
// claramente relacionado. As demais ficam soltas (sem parentesco óbvio).
$pais = ['Acidente', 'Manutenção', 'Falta ou Atraso de Colaborador', 'Reclamação', 'Reprova', 'Demora'];
$paiIds = [];
foreach ($pais as $nome) {
    $cat = new ITILCategory();
    $id = $cat->add([
        'name' => $nome,
        'entities_id' => ENTIDADE_OCORRENCIAS,
        'is_recursive' => 0,
        'is_helpdeskvisible' => 1,
        'is_incident' => 1,
        'is_request' => 1,
        'is_problem' => 1,
        'is_change' => 1,
    ]);
    $paiIds[$nome] = $id;
    echo "Categoria-pai '$nome' criada (id $id)\n";
}

$filhosPorNome = [
    'Acidente' => ['Acidente de Trabalho', 'Acidente de Trânsito', 'Acidente no Transporte'],
    'Manutenção' => ['Manutenção Elétrica', 'Manutenção Mecânica', 'Manutenção Predial'],
    'Falta ou Atraso de Colaborador' => ['Falta ou Atraso de Colaborador Terceiro', 'Falta ou Atraso de Colaborador TQUIM'],
    'Reclamação' => ['Reclamação da Sociedade', 'Reclamação de Cliente', 'Reclamação Interna'],
    'Reprova' => ['Reprova Interna', 'Reprova no Cliente'],
    'Demora' => ['Demora na adequação do tanque', 'Demora no carga ou descarga'],
];

foreach ($filhosPorNome as $pai => $filhos) {
    foreach ($filhos as $nomeFilho) {
        $it = $DB->request(['FROM' => 'glpi_itilcategories', 'WHERE' => ['name' => $nomeFilho, 'entities_id' => ENTIDADE_OCORRENCIAS]]);
        foreach ($it as $row) {
            $cat = new ITILCategory();
            $cat->update(['id' => $row['id'], 'itilcategories_id' => $paiIds[$pai]]);
            echo "  '$nomeFilho' -> pai '$pai'\n";
        }
    }
}

// 2. Grupos por unidade (quem abre a ocorrência escolhe manualmente).
foreach (['Ocorrências Armazém', 'Ocorrências SJP'] as $nome) {
    $grp = new Group();
    $id = $grp->add([
        'name' => $nome,
        'entities_id' => ENTIDADE_OCORRENCIAS,
        'is_recursive' => 1,
        'is_requester' => 1,
        'is_watcher' => 1,
        'is_assign' => 1,
        'is_task' => 1,
        'is_notify' => 1,
        'is_itemgroup' => 1,
        'is_usergroup' => 1,
        'is_manager' => 1,
    ]);
    echo "Grupo '$nome' criado (id $id)\n";
}

// 3. Bloco "Classificação da Qualidade": campo com os 14 códigos do Quadro 1
// (numeração igual à planilha da Qualidade) + move pra cá o campo "Inserir
// na avaliação motorista?", que sai da abertura do chamado e fica oculto de
// quem abre, só visível/editável nesta aba.
$field = new PluginFieldsField();
$fieldId = $field->add([
    'name' => 'codigoqualidadefield',
    'label' => 'Código da Ocorrência (Quadro 1 - Qualidade)',
    'type' => 'dropdown',
    'plugin_fields_containers_id' => CONTAINER_CLASSIFICACAO,
    'ranking' => 1,
    'is_active' => 1,
    'mandatory' => 0,
]);
echo "Campo 'codigoqualidadefield' criado (id $fieldId)\n";

$codigos = [
    '1 - Estação de Limpeza e ETE',
    '2 - Manutenção Veículos e Tanques',
    '3 - Programação do transporte',
    '4 - Comportamento ou condição inadequada/insegura (TQUIM)',
    '5 - Comportamento ou condição inadequada/insegura (Terceiros)',
    '6 - Cliente, demora na carga, na descarga, etc',
    '7 - Expedição e Faturamento de Doctos',
    '8 - Rastreador',
    '9 - Manutenção Predial, máquinas e equipamentos de produção',
    '10 - Cadastros e Gestão de Documentos e Inspeções',
    '11 - Recursos de TI, Comunicação, SEFAZ e Energia',
    '12 - Problemas de saúde',
    '13 - Condições adversas na rota/pista',
    '14 - Outros',
];
$dropdownClass = new PluginFieldsCodigoqualidadefieldDropdown();
foreach ($codigos as $nome) {
    $id = $dropdownClass->add(['name' => $nome]);
    echo "  código '$nome' (id $id)\n";
}

$campoAvaliacao = $DB->request(['FROM' => 'glpi_plugin_fields_fields', 'WHERE' => ['name' => 'inserirnaavaliaomotoristafield']]);
foreach ($campoAvaliacao as $row) {
    $field = new PluginFieldsField();
    $field->update(['id' => $row['id'], 'plugin_fields_containers_id' => CONTAINER_CLASSIFICACAO, 'ranking' => 2]);
    echo "Campo 'Inserir na avaliação motorista?' movido pro bloco de Classificação\n";
}

// 4. Data de abertura automática (igual ao template de TI): valor
// predefinido "NOW" faz o campo vir preenchido e travado (somente leitura).
$predef = new TicketTemplatePredefinedField();
$predef->add(['tickettemplates_id' => TEMPLATE_OCORRENCIAS, 'num' => 15, 'value' => 'NOW']);
echo "Data de abertura predefinida como NOW no template de Ocorrências\n";

echo "\nConcluído.\n";
