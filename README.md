# GLPI - TQUIM

Scripts de administração e relatórios para a instância GLPI da TQUIM Transportes, usados via API REST (`api.php/v1`).

Nenhum arquivo neste repositório contém credenciais, URLs de servidor real ou dados de colaboradores/clientes. Cada projeto tem um `config_exemplo.json` — copie para `config.json`, preencha com seus próprios dados e mantenha esse arquivo fora do controle de versão (já ignorado via `.gitignore`).

## Projetos

### [`atribuidor-entidade-usuarios/`](atribuidor-entidade-usuarios/)
Move o vínculo perfil+entidade (`Profile_User`) de um conjunto de usuários de uma entidade de origem para uma entidade de destino, mantendo perfil e recursividade. Modo simulação por padrão; exige confirmação explícita para aplicar.

### [`relatorio-ocorrencias/`](relatorio-ocorrencias/)
Exporta um relatório em Excel (`.xlsx`) formatado com os chamados de uma entidade, incluindo campos nativos e campos personalizados do plugin Fields, usando a API de busca do GLPI.

## Requisitos

- Python 3 (biblioteca padrão apenas no `atribuidor-entidade-usuarios`; `openpyxl` no `relatorio-ocorrencias`)
- Um cliente de API habilitado em **Configurar > Geral > API** no GLPI, e um token de API pessoal em **Preferências > Chaves de API remota**
