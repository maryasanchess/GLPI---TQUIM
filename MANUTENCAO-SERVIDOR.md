# Manutenção do servidor

Anotações operacionais do servidor que hospeda o GLPI. Nada aqui é específico
de um script — é sobre a máquina.

## Interface gráfica: desligar e religar

A VM roda com ambiente gráfico, que consome perto de 1 GB de RAM sem
contrapartida quando o acesso é por terminal. Desligar a interface é a forma
mais rápida de recuperar memória, sem mexer em hardware.

Desligar (reinicia em seguida):

```bash
sudo systemctl set-default multi-user.target && sudo reboot
```

Religar:

```bash
sudo systemctl set-default graphical.target && sudo reboot
```

Depois de desligar, o console do hipervisor passa a mostrar um login em texto
em vez da tela gráfica — o acesso continua igual, só muda a aparência.

**O GLPI não é afetado.** Ele é servido pelo Apache, como serviço; nunca
dependeu da área de trabalho. Quem acessa pelo navegador não percebe diferença.

## Memória e swap

Sintoma de falta de memória: a máquina fica lenta de forma generalizada — a
própria tela de login demora, o GLPI arrasta e buscas pesadas parecem travar.

Diagnóstico:

```bash
free -h; swapon --show; ps aux --sort=-%mem | head -8
```

O número que importa é o **swap em uso**. Swap é memória emprestada do disco:
se houver muita coisa nele, cada acesso a esses dados vira leitura de disco, que
é muito mais lento que RAM.

Em 27/08/2026, a VM tinha 3,3 GB de RAM com **3,0 GB de swap em uso** — ou seja,
precisava de algo perto de 6 GB e tinha metade disso. Para GLPI com banco na
mesma máquina, 4 GB é o mínimo confortável e 8 GB é folgado.

Ordem recomendada quando for ajustar:

1. desligar a interface gráfica (comando acima) e reiniciar;
2. conferir que o console em texto aparece e que o GLPI responde no navegador;
3. desligar a VM (`sudo poweroff`);
4. aumentar a memória nas configurações da VM, no hipervisor;
5. ligar e conferir com `free -h` — o swap em uso deve cair perto de zero.

**Não aumente o `innodb_buffer_pool_size` do MySQL antes de aumentar a RAM.**
Ampliar o buffer com a memória já estourada só aumenta o swap e piora a
lentidão. Depois de resolver a RAM, o valor de referência é 50–70% da memória
da VM.

## Atualizar um script no servidor

A pasta onde os scripts rodam **não é um clone deste repositório** — os arquivos
foram copiados manualmente. Por isso `git pull` não funciona lá, e o servidor
pode ficar defasado sem ninguém perceber.

Enquanto for assim, atualize baixando o arquivo direto do repositório. Exemplo
com o exportador de ocorrências:

```bash
cd /home/glpi/relatorio_ocorrencias && cp exportar_relatorio_ocorrencias.py exportar_relatorio_ocorrencias.py.bak && curl -fsSL https://raw.githubusercontent.com/maryasanchess/GLPI---TQUIM/main/relatorio-ocorrencias/exportar_relatorio_ocorrencias.py -o exportar_relatorio_ocorrencias.py && sha256sum exportar_relatorio_ocorrencias.py && python3 -c "import ast;ast.parse(open('exportar_relatorio_ocorrencias.py').read());print('sintaxe OK')"
```

O `.bak` é feito antes de sobrescrever, então o rollback é imediato:

```bash
cd /home/glpi/relatorio_ocorrencias && mv exportar_relatorio_ocorrencias.py.bak exportar_relatorio_ocorrencias.py
```

Para achar onde um script está instalado:

```bash
find / -name exportar_relatorio_ocorrencias.py -not -path '/proc/*' 2>/dev/null
```

## Agendamentos

Os scripts agendados rodam no crontab do **www-data**, não no de um usuário
comum. O motivo está no `relatorio-ocorrencias/LEIA-ME.txt`: cron não tem
terminal, então `sudo -u www-data` a partir do crontab de outro usuário falha
sempre — e, sem serviço de e-mail local, esse erro some sem deixar rastro.

```bash
sudo crontab -u www-data -l
```
