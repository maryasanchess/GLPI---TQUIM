<?php
// Envia o relatório de Ocorrências por e-mail reaproveitando a configuração
// de e-mail já existente no GLPI (mesmo mecanismo das notificações).
//
// Uso: php enviar_relatorio.php <caminho_do_arquivo.xlsx>

require_once '/var/www/html/glpi/src/Glpi/Application/ResourcesChecker.php';
require_once '/var/www/html/glpi/vendor/autoload.php';

use Glpi\Kernel\Kernel;

$kernel = new Kernel();
$kernel->boot();

if ($argc < 2) {
    fwrite(STDERR, "Uso: php enviar_relatorio.php <caminho_do_arquivo.xlsx>\n");
    exit(1);
}
$arquivo = $argv[1];
if (!is_file($arquivo)) {
    fwrite(STDERR, "Arquivo não encontrado: $arquivo\n");
    exit(1);
}

// Destinatários do relatório. Adicione mais linhas aqui conforme necessário.
$destinatarios = [
    ['email' => 'marya.souza@tquim.com.br', 'name' => 'Marya Souza'],
];

$mailer = new GLPIMailer();
foreach ($destinatarios as $dest) {
    $mailer->addAddress($dest['email'], $dest['name']);
}
$mailer->isHTML(true);

$email = $mailer->getEmail();
$email->subject('Relatório de Ocorrências - ' . basename($arquivo));
$email->html(
    '<p>Segue em anexo o relatório de chamados de Ocorrências, com uma aba anual '
    . 'e uma aba por mês.</p><p><em>Mensagem gerada automaticamente pelo sistema.</em></p>'
);
$email->attachFromPath($arquivo, basename($arquivo));

if (!$mailer->send()) {
    fwrite(STDERR, "Erro ao enviar: " . $mailer->getError() . "\n");
    exit(1);
}

echo "E-mail enviado com sucesso.\n";
