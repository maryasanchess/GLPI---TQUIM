<?php
// Verifica se chegou um e-mail pedindo o relatório de Ocorrências fora do
// dia agendado. Roda em polling (via cron), lendo a MESMA caixa que o GLPI
// já usa pra abrir chamados (chamados.ti@tquim.com.br, MailCollector id 1)
// - reaproveita a credencial já cadastrada no GLPI, sem precisar guardar
// senha de e-mail em lugar nenhum deste script.
//
// COMO PEDIR: mandar um e-mail pra chamados.ti@tquim.com.br com o assunto
// contendo a palavra "RELATORIO OCORRENCIAS" (não sensível a maiúsculas),
// de um endereço autorizado (lista $remetentesAutorizados abaixo). O
// relatório do ano corrente é gerado e enviado de volta pro remetente.
//
// Não apaga nenhum e-mail: só marca como lido depois de processar (e só
// olha e-mails NÃO lidos a cada execução), pra não reprocessar o mesmo
// pedido, sem apagar nada da caixa de forma automática/não supervisionada.
// Os e-mails de pedido vão se acumular (lidos) na caixa até alguém limpar
// manualmente de vez em quando.
//
// RISCO ACEITO (decisão consciente, não é bug): como essa é a MESMA caixa
// que o coletor de e-mail nativo do GLPI usa pra abrir chamados, existe uma
// chance pequena do GLPI processar o e-mail-gatilho antes deste script e
// abrir um chamado indevido com ele (esse chamado precisaria ser apagado
// manualmente se acontecer). Pra minimizar, agende este script pra rodar
// com MAIS frequência que o coletor de e-mail do GLPI (que roda a cada 5
// minutos) - ver LEIA-ME.txt.
//
// Uso: php verificar_solicitacao_relatorio.php

require_once '/var/www/html/glpi/src/Glpi/Application/ResourcesChecker.php';
require_once '/var/www/html/glpi/vendor/autoload.php';

use Glpi\Kernel\Kernel;

$kernel = new Kernel();
$kernel->boot();

const MAILCOLLECTOR_ID = 1;
const PALAVRA_CHAVE = 'RELATORIO OCORRENCIAS';
const PASTA_SCRIPT = __DIR__;

$remetentesAutorizados = [
    'marya.souza@tquim.com.br',
];

$collector = new MailCollector();
if (!$collector->getFromDB(MAILCOLLECTOR_ID)) {
    fwrite(STDERR, "MailCollector id " . MAILCOLLECTOR_ID . " não encontrado.\n");
    exit(1);
}

$conexao = $collector->connect();
if (!$conexao) {
    fwrite(STDERR, "Não foi possível conectar na caixa de e-mail: " . imap_last_error() . "\n");
    exit(1);
}

// Só e-mails ainda não lidos - evita reprocessar o mesmo pedido de novo.
$mensagens = imap_search($conexao, 'UNSEEN');
$processados = 0;

if ($mensagens !== false) {
    foreach ($mensagens as $numero) {
        $header = imap_headerinfo($conexao, $numero);
        if (!$header) {
            continue;
        }
        $assunto = isset($header->subject) ? imap_utf8($header->subject) : '';
        if (stripos($assunto, PALAVRA_CHAVE) === false) {
            continue; // não é um pedido de relatório - deixa não lido, sem mexer
        }

        $remetente = strtolower(trim(($header->from[0]->mailbox ?? '') . '@' . ($header->from[0]->host ?? '')));
        $autorizado = in_array($remetente, array_map('strtolower', $remetentesAutorizados), true);

        echo "E-mail encontrado: assunto='$assunto' de='$remetente' autorizado=" . ($autorizado ? 'sim' : 'não') . "\n";

        if ($autorizado) {
            $ano = date('Y');
            $arquivo = PASTA_SCRIPT . "/Relatorio_Ocorrencias_{$ano}_sob_demanda.xlsx";

            $comando = 'cd ' . escapeshellarg(PASTA_SCRIPT)
                . ' && /usr/bin/python3 exportar_relatorio_ocorrencias.py '
                . escapeshellarg($arquivo) . ' ' . escapeshellarg((string) $ano) . ' 2>&1';
            exec($comando, $saidaComando, $codigoRetorno);
            echo implode("\n", $saidaComando) . "\n";

            if ($codigoRetorno === 0 && is_file($arquivo)) {
                $config = Config::getConfigurationValues('core', ['admin_email', 'admin_email_name']);
                $mailer = new GLPIMailer();
                $mailer->setFrom($config['admin_email'], $config['admin_email_name'] ?? '');
                $mailer->addAddress($remetente);
                $mailer->isHTML(true);
                $email = $mailer->getEmail();
                $email->subject('Relatório de Ocorrências - solicitado por e-mail');
                $email->html(
                    '<p>Segue o relatório de Ocorrências de ' . $ano . ' solicitado por e-mail.</p>'
                    . '<p><em>Gerado automaticamente em resposta ao seu pedido.</em></p>'
                );
                $email->attachFromPath($arquivo, basename($arquivo));
                if ($mailer->send()) {
                    echo "Relatório enviado para $remetente.\n";
                } else {
                    fwrite(STDERR, "Erro ao enviar: " . $mailer->getError() . "\n");
                }
            } else {
                fwrite(STDERR, "Erro ao gerar o relatório (código $codigoRetorno).\n");
            }
        }

        // Marca como lido (não apaga nada) - não é processado de novo na
        // próxima execução, mesmo que não fosse de um remetente autorizado.
        imap_setflag_full($conexao, (string) $numero, '\\Seen');
        $processados++;
    }
}

imap_close($conexao);

echo "Concluído. $processados e-mail(s) de solicitação processado(s).\n";
