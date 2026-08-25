<?php
// Cria um novo dashboard "Ocorrências" reaproveitando os cards do dashboard
// padrão "Assistência" do GLPI (Dashboard::getDefaults()['assistance']).
//
// Uso: php criar_dashboard_ocorrencias.php

require_once '/var/www/html/glpi/src/Glpi/Application/ResourcesChecker.php';
require_once '/var/www/html/glpi/vendor/autoload.php';

use Glpi\Kernel\Kernel;
use Glpi\Dashboard\Dashboard;

$kernel = new Kernel();
$kernel->boot();

global $DB;

$defaults = Dashboard::getDefaults();
if (!isset($defaults['assistance'])) {
    fwrite(STDERR, "Preset 'assistance' não encontrado.\n");
    exit(1);
}
$preset = $defaults['assistance'];

$dashboardKey = 'ocorrencias';

$existe = $DB->request(['FROM' => 'glpi_dashboards_dashboards', 'WHERE' => ['key' => $dashboardKey]]);
foreach ($existe as $row) {
    fwrite(STDERR, "Dashboard '$dashboardKey' já existe (id {$row['id']}). Nada foi feito.\n");
    exit(1);
}

$DB->insert('glpi_dashboards_dashboards', [
    'key'      => $dashboardKey,
    'name'     => 'Ocorrências',
    'context'  => 'core',
    'users_id' => 0,
]);
$dashboardId = $DB->insertId();

echo "Dashboard '$dashboardKey' criado (id $dashboardId).\n";

$inseridos = 0;
foreach ($preset['items'] as $item) {
    $cardId = $item['card_id'];
    $uuid = bin2hex(random_bytes(16));
    $gridstackId = $cardId . '_' . substr($uuid, 0, 8) . '-' . substr($uuid, 8, 4) . '-'
        . substr($uuid, 12, 4) . '-' . substr($uuid, 16, 4) . '-' . substr($uuid, 20, 12);

    $DB->insert('glpi_dashboards_items', [
        'dashboards_dashboards_id' => $dashboardId,
        'gridstack_id'             => $gridstackId,
        'card_id'                  => $cardId,
        'x'                        => $item['x'],
        'y'                        => $item['y'],
        'width'                    => $item['width'],
        'height'                   => $item['height'],
        'card_options'             => json_encode($item['card_options'] ?? []),
    ]);
    $inseridos++;
}

echo "$inseridos card(s) copiado(s) do modelo 'Assistência'.\n";
echo "Acesse em: Central > Dashboards, ou pela URL /front/central.php?dashboard=$dashboardKey\n";
echo "IMPORTANTE: para ver só os dados de Ocorrências, troque a entidade ativa para TQUIM > Ocorrências antes de abrir.\n";
