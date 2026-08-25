<?php
// Notifica por e-mail, com os valores reais dos 37 campos preenchidos, sempre
// que um novo chamado é criado na entidade Ocorrências. Roda periodicamente
// (via cron, a cada poucos minutos) e envia só o que ainda não foi enviado,
// controlado por um arquivo de estado local (ultimo_id.txt).
//
// Substitui a notificação nativa do GLPI para "Novo chamado" na entidade
// Ocorrências (que não consegue mostrar os campos do plugin Fields).
//
// Uso: php notificar_ocorrencia.php

require_once '/var/www/html/glpi/src/Glpi/Application/ResourcesChecker.php';
require_once '/var/www/html/glpi/vendor/autoload.php';

use Glpi\Kernel\Kernel;

$kernel = new Kernel();
$kernel->boot();

const ENTIDADE_OCORRENCIAS = 6;
const ARQUIVO_ESTADO = __DIR__ . '/ultimo_id.txt';

$destinatarios = [
    ['email' => 'marya.souza@tquim.com.br', 'name' => 'Marya Souza'],
];

$yesno = static fn($v) => ($v === '' || $v === null) ? '' : (((int) $v === 1) ? 'Sim' : 'Não');

function campo(string $coluna): callable
{
    return static fn(array $d) => (string) ($d[$coluna] ?? '');
}

function campoFmt(string $coluna, callable $fmt): callable
{
    return static function (array $d) use ($coluna, $fmt) {
        $v = $d[$coluna] ?? '';
        return ($v === '' || $v === null) ? '' : $fmt($v);
    };
}

function dropdown(string $colunaFk, string $tabela): callable
{
    return static function (array $d) use ($colunaFk, $tabela) {
        return resolverDropdown((int) ($d[$colunaFk] ?? 0), $tabela);
    };
}

function resolverDropdown(int $id, string $tabela): string
{
    global $DB;
    if ($id <= 0) {
        return '';
    }
    $it = $DB->request(['SELECT' => 'name', 'FROM' => $tabela, 'WHERE' => ['id' => $id]]);
    foreach ($it as $row) {
        return (string) $row['name'];
    }
    return '';
}

// Layout na mesma ordem/agrupamento do modelo em .docx (Comunicado de
// Ocorrência), seguido do bloco de Tratativa/Classificação (só existe na
// planilha de controle, sem equivalente no .docx original).
// Cada item: ['unico', rótulo, getter] ou ['par', rótulo1, getter1, rótulo2, getter2]
$layout = [
    ['par', 'Placa Tração', campo('placatraofield'), 'Cód. Frota Tração', campo('cdfrotatraofield')],
    ['par', 'Placa Semi-Reboque', campo('placasemireboquefield'), 'Cód. Frota Semi Reboque', campo('cdfrotasemireboquefield')],
    ['unico', 'Motorista/Colaborador', campo('colaboradormotoristafield')],
    ['par', 'Cargo/Função', campo('cargofunofield'), 'Depto/Setor', campo('deptosetorfield')],
    ['par', 'OC / CT-e / NF', campo('occtefield'), 'Situação da Carga', dropdown('plugin_fields_situaodacargafielddropdowns_id', 'glpi_plugin_fields_situaodacargafielddropdowns')],
    ['unico', 'Produto', campo('produtofield')],
    ['unico', 'Cliente', campo('clientefield')],
    ['unico', 'Origem', campo('origemfield')],
    ['unico', 'Destino', campo('destinofield')],
    ['unico', 'Local da Ocorrência', campo('localdaocorrnciafield')],
    ['unico', 'Descrição da Ocorrência', campo('descriodaocorrnciafield')],
    ['unico', 'Providências já Tomadas', campo('providnciasjtomadafield')],
    ['unico', 'Responsável pela Análise / Ações Corretivas', campo('responsvelpelaanliseaescorretivafield')],
    ['par', 'Início da Jornada', campo('inciodajornadafield'), 'Fim da Jornada', campo('fimdajornadafield')],
    ['par', 'Início do Carregamento', campo('inciodocarregamentofield'), 'Fim do Carregamento', campo('fimdocarregamentofield')],
    ['par', 'Motivo', dropdown('plugin_fields_motivofielddropdowns_id', 'glpi_plugin_fields_motivofielddropdowns'), 'Responsável', dropdown('plugin_fields_responsvelclientemotoristafielddropdowns_id', 'glpi_plugin_fields_responsvelclientemotoristafielddropdowns')],
    ['unico', 'Impacto ao Cliente?', campoFmt('impactoaoclientefield', $yesno)],
    ['unico', 'Outras Observações', campo('outrasobservaefield')],
    ['unico', 'Lançar na avaliação do colaborador?', campoFmt('inserirnaavaliaomotoristafield', $yesno)],
    ['unico', 'Custos da NC', campo('custosdancfield')],
];

$layoutTratativa = [
    ['par', 'Cod. Multa', campo('codmultafield'), 'S.A.', campo('safield')],
    ['unico', 'Plano de Ações', campo('planodeaefield')],
    ['unico', 'Levantamento de Custos', campo('levantamentodecustofield')],
    ['par', 'Quantidade que vazou', campo('quantidadequevazoufield'), 'Custo Total', campo('custototalfield')],
    ['par', 'Acionamento SuatransPamcary?', campoFmt('acionamentosuatranspamcaryfield', $yesno), 'Houve vazamento?', campoFmt('houvevazamentofield', $yesno)],
    ['par', 'Tipo de NC', dropdown('plugin_fields_tipodencfielddropdowns_id', 'glpi_plugin_fields_tipodencfielddropdowns'), 'Classificação', dropdown('plugin_fields_classificaofielddropdowns_id', 'glpi_plugin_fields_classificaofielddropdowns')],
];

function montarLinhaUnica(string $rotulo, string $valor): string
{
    $v = htmlspecialchars($valor !== '' ? $valor : '-', ENT_QUOTES, 'UTF-8');
    $r = htmlspecialchars($rotulo, ENT_QUOTES, 'UTF-8');
    return '<tr>'
        . '<td style="background: #f8f9fb; padding: 8px; border-bottom: 1px solid #e5e5e5; border-right: 1px solid #e5e5e5;" width="220"><strong>' . $r . '</strong></td>'
        . '<td style="padding: 8px; border-bottom: 1px solid #e5e5e5;" colspan="3">' . $v . '</td>'
        . '</tr>';
}

function montarLinhaPar(string $r1, string $v1, string $r2, string $v2): string
{
    $r1 = htmlspecialchars($r1, ENT_QUOTES, 'UTF-8');
    $r2 = htmlspecialchars($r2, ENT_QUOTES, 'UTF-8');
    $v1 = htmlspecialchars($v1 !== '' ? $v1 : '-', ENT_QUOTES, 'UTF-8');
    $v2 = htmlspecialchars($v2 !== '' ? $v2 : '-', ENT_QUOTES, 'UTF-8');
    return '<tr>'
        . '<td style="background: #f8f9fb; padding: 8px; border-bottom: 1px solid #e5e5e5; border-right: 1px solid #e5e5e5;" width="180"><strong>' . $r1 . '</strong></td>'
        . '<td style="padding: 8px; border-bottom: 1px solid #e5e5e5; border-right: 1px solid #e5e5e5;" width="170">' . $v1 . '</td>'
        . '<td style="background: #f8f9fb; padding: 8px; border-bottom: 1px solid #e5e5e5; border-right: 1px solid #e5e5e5;" width="180"><strong>' . $r2 . '</strong></td>'
        . '<td style="padding: 8px; border-bottom: 1px solid #e5e5e5;">' . $v2 . '</td>'
        . '</tr>';
}

function montarLinhas(array $layout, array $dados): string
{
    $out = '';
    foreach ($layout as $item) {
        if ($item[0] === 'unico') {
            [, $rotulo, $getter] = $item;
            $out .= montarLinhaUnica($rotulo, $getter($dados));
        } else {
            [, $r1, $g1, $r2, $g2] = $item;
            $out .= montarLinhaPar($r1, $g1($dados), $r2, $g2($dados));
        }
    }
    return $out;
}

function montarTituloSecao(string $texto): string
{
    $t = htmlspecialchars($texto, ENT_QUOTES, 'UTF-8');
    return '<tr><td colspan="4" style="background: #FFFF99; color: #333333; padding: 8px 10px; font-weight: bold; border-bottom: 1px solid #d4d9e2;">' . $t . '</td></tr>';
}

function montarEmail(array $ticket, array $dados, array $layout, array $layoutTratativa): string
{
    $linhas = montarLinhaPar('Categoria', $ticket['categoria'], 'Prioridade', $ticket['prioridade']);
    $linhas .= montarLinhas($layout, $dados);
    $linhas .= montarTituloSecao('TRATATIVA E CLASSIFICAÇÃO');
    $linhas .= montarLinhas($layoutTratativa, $dados);

    $url = rtrim($GLOBALS['CFG_GLPI']['url_base'] ?? '', '/') . '/front/ticket.form.php?id=' . $ticket['id'];

    return '<table style="background: #e9edf5; width: 100%; font-family: Arial,Helvetica,sans-serif;" border="0" cellspacing="0" cellpadding="0">'
        . '<tbody><tr><td align="center">'
        . '<table style="background: #ffffff; width: 700px; border: 1px solid #d4d9e2;" border="0" cellspacing="0" cellpadding="0">'
        . '<tbody>'
        . '<tr><td style="padding: 25px; text-align: center; background: #ffffff;">'
        . '<img src="https://www.tquim.com.br/wp-content/uploads/2019/05/logo-tquim-retina.png" alt="TQUIM" width="180"></td></tr>'
        . '<tr><td style="background: #FFFF99; color: #333333; padding: 15px; text-align: center; font-size: 20px; font-weight: bold; border-bottom: 1px solid #d4d9e2;">COMUNICADO DE OCORRÊNCIA</td></tr>'
        . '<tr><td style="background: #fff8d6; padding: 10px 20px; text-align: center; font-size: 12px; font-style: italic; color: #555555;">'
        . 'O conteúdo desta ocorrência é confidencial e deve ficar restrito à TQUIM. É proibido compartilhar informações e imagens fora da empresa.</td></tr>'
        . '<tr><td style="background: #FFFF99; padding: 16px 20px; text-align: center; font-size: 19px; font-weight: bold; color: #333333; border-bottom: 1px solid #d4d9e2;">'
        . htmlspecialchars($ticket['titulo'], ENT_QUOTES, 'UTF-8') . '</td></tr>'
        . '<tr><td style="background: #f3f5f8; padding: 12px; text-align: center; border-bottom: 1px solid #d8d8d8;"><strong>Chamado Nº '
        . str_pad((string) $ticket['id'], 7, '0', STR_PAD_LEFT) . '</strong></td></tr>'
        . '<tr><td style="padding: 25px;">'
        . '<table style="border-collapse: collapse;" width="100%" cellspacing="0" cellpadding="0">'
        . '<tbody>' . $linhas . '</tbody></table>'
        . '<div style="margin-top: 25px; text-align: center;">'
        . '<a href="' . $url . '" style="background: #0a67d8; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">VER CHAMADO COMPLETO</a>'
        . '</div></td></tr>'
        . '<tr><td style="background: #f3f5f8; padding: 15px; text-align: center; font-size: 12px; color: #888888;">Mensagem gerada automaticamente pelo sistema — TQUIM Transportes</td></tr>'
        . '</tbody></table></td></tr></tbody></table>';
}

global $DB;

$ultimoId = 0;
if (is_file(ARQUIVO_ESTADO)) {
    $ultimoId = (int) trim(file_get_contents(ARQUIVO_ESTADO));
}

$criteriosTickets = [
    'SELECT' => ['t.id', 't.name AS titulo', 't.priority', 'c.name AS categoria'],
    'FROM' => 'glpi_tickets AS t',
    'LEFT JOIN' => [
        'glpi_itilcategories AS c' => [
            'FKEY' => ['c' => 'id', 't' => 'itilcategories_id'],
        ],
    ],
    'WHERE' => [
        't.entities_id' => ENTIDADE_OCORRENCIAS,
        't.is_deleted' => 0,
        't.id' => ['>', $ultimoId],
    ],
    'ORDER' => 't.id ASC',
];

$prioridades = [1 => 'Muito baixa', 2 => 'Baixa', 3 => 'Média', 4 => 'Alta', 5 => 'Muito alta', 6 => 'Crítica'];

$maiorId = $ultimoId;
$enviados = 0;

foreach ($DB->request($criteriosTickets) as $ticket) {
    $maiorId = max($maiorId, (int) $ticket['id']);

    $dadosIt = $DB->request([
        'SELECT' => '*',
        'FROM' => 'glpi_plugin_fields_ticketdadosdaocorrncias',
        'WHERE' => ['items_id' => $ticket['id'], 'itemtype' => 'Ticket'],
    ]);
    $dados = [];
    foreach ($dadosIt as $row) {
        $dados = $row;
        break;
    }

    $ticketInfo = [
        'id' => $ticket['id'],
        'titulo' => $ticket['titulo'],
        'categoria' => $ticket['categoria'] ?? '',
        'prioridade' => $prioridades[(int) $ticket['priority']] ?? '',
    ];

    $html = montarEmail($ticketInfo, $dados, $layout, $layoutTratativa);

    $config = Config::getConfigurationValues('core', ['admin_email', 'admin_email_name']);
    $mailer = new GLPIMailer();
    $mailer->setFrom($config['admin_email'], $config['admin_email_name'] ?? '');
    foreach ($destinatarios as $dest) {
        $mailer->addAddress($dest['email'], $dest['name']);
    }
    $mailer->isHTML(true);
    $email = $mailer->getEmail();
    $email->subject('COMUNICADO DE OCORRÊNCIA - ' . $ticket['titulo']);
    $email->html($html);

    if ($mailer->send()) {
        echo "Enviado: chamado {$ticket['id']}\n";
        $enviados++;
    } else {
        fwrite(STDERR, "Erro ao enviar chamado {$ticket['id']}: " . $mailer->getError() . "\n");
    }
}

file_put_contents(ARQUIVO_ESTADO, (string) $maiorId);
echo "Concluído. {$enviados} e-mail(s) enviado(s). Último ID processado: {$maiorId}.\n";
