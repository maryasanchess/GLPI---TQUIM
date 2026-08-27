# GLPI - TQUIM

Scripts de administração e relatórios para a instância GLPI da TQUIM Transportes, usados via API REST (`api.php/v1`).

Nenhum arquivo neste repositório contém credenciais nem dados de chamados. Cada projeto tem um `config_exemplo.json` — copie para `config.json`, preencha com seus próprios dados e mantenha esse arquivo fora do controle de versão (já ignorado via `.gitignore`).

Alguns scripts de notificação trazem endereços de e-mail internos da TQUIM nas listas de destinatários, e a documentação de manutenção cita caminhos do servidor. Ao reaproveitar em outro ambiente, troque esses valores.

## Projetos

### [`atribuidor-entidade-usuarios/`](atribuidor-entidade-usuarios/)
Move o vínculo perfil+entidade (`Profile_User`) de um conjunto de usuários de uma entidade de origem para uma entidade de destino, mantendo perfil e recursividade. Modo simulação por padrão; exige confirmação explícita para aplicar.

### [`relatorio-ocorrencias/`](relatorio-ocorrencias/)
Exporta um relatório em Excel (`.xlsx`) formatado com os chamados de uma entidade, incluindo campos nativos e campos personalizados do plugin Fields. Traz um painel anual interativo, uma aba por mês e a notificação de abertura por e-mail com os valores reais dos campos do plugin.

### [`importador-chamados/`](importador-chamados/)
Lê a planilha de aberturas de chamado e cria/atualiza os chamados correspondentes no GLPI via API, sem disparar notificação para os requerentes. Controla o vínculo linha da planilha → chamado para não duplicar em reexecuções.

## Operação

### [`MANUTENCAO-SERVIDOR.md`](MANUTENCAO-SERVIDOR.md)
Anotações do servidor: como atualizar um script instalado, diagnóstico de memória/swap, como desligar e religar a interface gráfica, e onde ficam os agendamentos.

## Requisitos

- Python 3 (biblioteca padrão apenas no `atribuidor-entidade-usuarios`; `openpyxl` no `relatorio-ocorrencias`)
- Um cliente de API habilitado em **Configurar > Geral > API** no GLPI, e um token de API pessoal em **Preferências > Chaves de API remota**
