<?php
// Permite pedir o relatório de Ocorrências por e-mail, fora do dia
// agendado do mês.
//
// COMO PEDIR: mandar um e-mail pra chamados.ti@tquim.com.br com
// "RELATORIO OCORRENCIAS" no assunto, de um endereço autorizado (lista
// $remetentesAutorizados abaixo).
//
// COMO FUNCIONA: NÃO lê a caixa de e-mail diretamente (tentamos isso
// primeiro com a extensão imap do PHP, mas essa versão do GLPI usa uma
// biblioteca PHP pura pra e-mail, incompatível com as funções imap_*
// nativas, e o pacote do sistema pra isso nem tinha candidato de
// instalação). Em vez disso, deixa o PRÓPRIO coletor de e-mail nativo do
// GLPI (já testado, já funciona) transformar o e-mail-gatilho num chamado
// normal, e este script roda em polling (via cron, mesmo padrão do
// notificar_ocorrencia.php) procurando chamados novos, em QUALQUER
// entidade, com esse assunto. Ao achar um chamado autorizado: gera e manda
// o relatório pro requerente, depois fecha o chamado com um comentário.
// Chamados de remetente não autorizado ficam abertos, sem resposta
// automática, pra alguém da TI olhar manualmente.
//
// Uso: php verificar_solicitacao_relatorio.php

require_once '/var/www/html/glpi/src/Glpi/Application/ResourcesChecker.php';
require_once '/var/www/html/glpi/vendor/autoload.php';

use Glpi\Kernel\Kernel;

$kernel = new Kernel();
$kernel->boot();

const PALAVRA_CHAVE = 'RELATORIO OCORRENCIAS';
const PASTA_SCRIPT = __DIR__;
const ARQUIVO_ESTADO = PASTA_SCRIPT . '/ultimo_id_solicitacao.txt';

$remetentesAutorizados = [
    'marya.souza@tquim.com.br',
];

global $DB;

$ultimoId = 0;
if (is_file(ARQUIVO_ESTADO)) {
    $ultimoId = (int) trim(file_get_contents(ARQUIVO_ESTADO));
}

$criterios = [
    'SELECT' => ['id', 'name'],
    'FROM' => 'glpi_tickets',
    'WHERE' => [
        'is_deleted' => 0,
        'id' => ['>', $ultimoId],
        'name' => ['LIKE', '%' . PALAVRA_CHAVE . '%'],
    ],
    'ORDER' => 'id ASC',
];

$maiorId = $ultimoId;
$processados = 0;

foreach ($DB->request($criterios) as $row) {
    $ticketId = (int) $row['id'];
    $maiorId = max($maiorId, $ticketId);

    $ticket = new Ticket();
    if (!$ticket->getFromDB($ticketId)) {
        continue;
    }

    $emailRequerente = '';
    $nomeRequerente = '';
    $requerentes = $ticket->getUsers(CommonITILActor::REQUESTER);
    foreach ($requerentes as $r) {
        $userId = (int) ($r['users_id'] ?? 0);
        if ($userId > 0) {
            $user = new User();
            if ($user->getFromDB($userId)) {
                $email = $user->getDefaultEmail();
                if ($email) {
                    $emailRequerente = strtolower($email);
                    $nomeRequerente = $user->getFriendlyName() ?: $emailRequerente;
                    break;
                }
            }
        }
    }

    $autorizado = $emailRequerente !== '' && in_array($emailRequerente, array_map('strtolower', $remetentesAutorizados), true);
    echo "Chamado {$ticketId}: requerente='{$emailRequerente}' autorizado=" . ($autorizado ? 'sim' : 'não') . "\n";

    if ($autorizado) {
        $ano = date('Y');
        $arquivo = PASTA_SCRIPT . "/Relatorio_Ocorrencias_{$ano}_sob_demanda.xlsx";

        $comando = 'cd ' . escapeshellarg(PASTA_SCRIPT)
            . ' && /usr/bin/python3 exportar_relatorio_ocorrencias.py '
            . escapeshellarg($arquivo) . ' ' . escapeshellarg((string) $ano) . ' 2>&1';
        exec($comando, $saidaComando, $codigoRetorno);
        echo implode("\n", $saidaComando) . "\n";

        $comentario = '';
        if ($codigoRetorno === 0 && is_file($arquivo)) {
            $config = Config::getConfigurationValues('core', ['admin_email', 'admin_email_name']);
            $mailer = new GLPIMailer();
            $mailer->setFrom($config['admin_email'], $config['admin_email_name'] ?? '');
            $mailer->addAddress($emailRequerente, $nomeRequerente);
            $mailer->isHTML(true);
            $email = $mailer->getEmail();
            $email->subject('Relatório de Ocorrências - solicitado por e-mail');
            $email->html(
                '<p>Segue o relatório de Ocorrências de ' . $ano . ' solicitado por e-mail.</p>'
                . '<p><em>Gerado automaticamente em resposta ao seu pedido.</em></p>'
            );
            $email->attachFromPath($arquivo, basename($arquivo));
            if ($mailer->send()) {
                echo "Relatório enviado para {$emailRequerente}.\n";
                $comentario = 'Relatório de Ocorrências gerado e enviado automaticamente para ' . $emailRequerente . '.';
            } else {
                fwrite(STDERR, "Erro ao enviar: " . $mailer->getError() . "\n");
                $comentario = 'Falha ao enviar o relatório: ' . $mailer->getError();
            }
        } else {
            fwrite(STDERR, "Erro ao gerar o relatório (código {$codigoRetorno}).\n");
            $comentario = 'Falha ao gerar o relatório (código ' . $codigoRetorno . ').';
        }

        // Fecha o chamado-gatilho automaticamente, com um comentário interno
        // explicando o que aconteceu.
        $followup = new ITILFollowup();
        $followup->add([
            'itemtype' => 'Ticket',
            'items_id' => $ticketId,
            'content' => $comentario,
            '_disablenotifications' => true,
        ]);
        // Recarrega o chamado do zero antes de fechar - reusar o objeto que
        // já rodou getUsers() nele pode não atualizar o status de verdade.
        $ticketFechar = new Ticket();
        $ticketFechar->getFromDB($ticketId);
        $fechou = $ticketFechar->update(['id' => $ticketId, 'status' => 6, '_disablenotifications' => true]);
        echo $fechou ? "Chamado {$ticketId} fechado.\n" : "AVISO: não consegui fechar o chamado {$ticketId}.\n";
    }

    $processados++;
}

file_put_contents(ARQUIVO_ESTADO, (string) $maiorId);
echo "Concluído. {$processados} pedido(s) de relatório processado(s). Último ID: {$maiorId}.\n";
