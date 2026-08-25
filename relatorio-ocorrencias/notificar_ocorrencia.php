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

$yesno = static fn($v) => ((int) $v === 1) ? 'Sim' : 'Não';

// (rótulo, coluna na tabela principal, formatador opcional)
$camposDiretos = [
    ['Placa Tração', 'placatraofield', null],
    ['Placa Semi-Reboque', 'placasemireboquefield', null],
    ['Cód. Frota Tração', 'cdfrotatraofield', null],
    ['Cód. Frota Semi Reboque', 'cdfrotasemireboquefield', null],
    ['Colaborador Motorista', 'colaboradormotoristafield', null],
    ['Cargo/Função', 'cargofunofield', null],
    ['Depto/Setor', 'deptosetorfield', null],
    ['OC / CT-e / NF', 'occtefield', null],
    ['Produto', 'produtofield', null],
    ['Origem', 'origemfield', null],
    ['Destino', 'destinofield', null],
    ['Cliente', 'clientefield', null],
    ['Local da Ocorrência', 'localdaocorrnciafield', null],
    ['Descrição da Ocorrência', 'descriodaocorrnciafield', null],
    ['Providências já Tomadas', 'providnciasjtomadafield', null],
    ['Responsável pela Análise / Ações Corretivas', 'responsvelpelaanliseaescorretivafield', null],
    ['Início da Jornada', 'inciodajornadafield', null],
    ['Fim da Jornada', 'fimdajornadafield', null],
    ['Início do Carregamento', 'inciodocarregamentofield', null],
    ['Fim do Carregamento', 'fimdocarregamentofield', null],
    ['Impacto ao Cliente', 'impactoaoclientefield', $yesno],
    ['Outras Observações', 'outrasobservaefield', null],
    ['Inserir na avaliação motorista', 'inserirnaavaliaomotoristafield', $yesno],
    ['Custos da NC', 'custosdancfield', null],
    ['Cod. Multa', 'codmultafield', null],
    ['S.A.', 'safield', null],
    ['Plano de Ações', 'planodeaefield', null],
    ['Levantamento de Custos', 'levantamentodecustofield', null],
    ['Quantidade que vazou', 'quantidadequevazoufield', null],
    ['Custo Total', 'custototalfield', null],
    ['Acionamento SuatransPamcary', 'acionamentosuatranspamcaryfield', $yesno],
    ['Houve vazamento', 'houvevazamentofield', $yesno],
];

// (rótulo, coluna da FK na tabela principal, tabela do dropdown)
$camposDropdown = [
    ['Situação da Carga', 'plugin_fields_situaodacargafielddropdowns_id', 'glpi_plugin_fields_situaodacargafielddropdowns'],
    ['Motivo', 'plugin_fields_motivofielddropdowns_id', 'glpi_plugin_fields_motivofielddropdowns'],
    ['Tipo de NC', 'plugin_fields_tipodencfielddropdowns_id', 'glpi_plugin_fields_tipodencfielddropdowns'],
    ['Classificação', 'plugin_fields_classificaofielddropdowns_id', 'glpi_plugin_fields_classificaofielddropdowns'],
    ['Responsável (Cliente/Motorista)', 'plugin_fields_responsvelclientemotoristafielddropdowns_id', 'glpi_plugin_fields_responsvelclientemotoristafielddropdowns'],
];

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

function montarLinhaTabela(string $rotulo, string $valor): string
{
    if ($valor === '' || $valor === null) {
        return '';
    }
    $valorEscapado = htmlspecialchars($valor, ENT_QUOTES, 'UTF-8');
    $rotuloEscapado = htmlspecialchars($rotulo, ENT_QUOTES, 'UTF-8');
    return '<tr><td style="background: #f8f9fb; padding: 8px; border-bottom: 1px solid #e5e5e5;" width="220"><strong>'
        . $rotuloEscapado . '</strong></td>'
        . '<td style="padding: 8px; border-bottom: 1px solid #e5e5e5;">' . $valorEscapado . '</td></tr>';
}

function montarEmail(array $ticket, array $dados, array $camposDiretos, array $camposDropdown): string
{
    $linhas = '';
    $linhas .= montarLinhaTabela('Categoria', $ticket['categoria']);
    $linhas .= montarLinhaTabela('Prioridade', $ticket['prioridade']);

    foreach ($camposDiretos as [$rotulo, $coluna, $formatador]) {
        $valor = $dados[$coluna] ?? '';
        if ($formatador !== null && $valor !== '' && $valor !== null) {
            $valor = $formatador($valor);
        }
        $linhas .= montarLinhaTabela($rotulo, (string) $valor);
    }
    foreach ($camposDropdown as [$rotulo, $colunaFk, $tabela]) {
        $valor = resolverDropdown((int) ($dados[$colunaFk] ?? 0), $tabela);
        $linhas .= montarLinhaTabela($rotulo, $valor);
    }

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

$sql = 'SELECT t.id, t.name AS titulo, t.priority, c.name AS categoria
        FROM glpi_tickets t
        LEFT JOIN glpi_itilcategories c ON c.id = t.itilcategories_id
        WHERE t.entities_id = ' . ((int) ENTIDADE_OCORRENCIAS) . '
          AND t.is_deleted = 0
          AND t.id > ' . ((int) $ultimoId) . '
        ORDER BY t.id ASC';

$prioridades = [1 => 'Muito baixa', 2 => 'Baixa', 3 => 'Média', 4 => 'Alta', 5 => 'Muito alta', 6 => 'Crítica'];

$maiorId = $ultimoId;
$enviados = 0;

foreach ($DB->request(['SQL' => $sql]) as $ticket) {
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

    $html = montarEmail($ticketInfo, $dados, $camposDiretos, $camposDropdown);

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
