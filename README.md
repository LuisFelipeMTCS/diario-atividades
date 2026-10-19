# Diário de Atividades

App local para anotar o que foi feito no dia e gerar o relatório mensal para o cliente.

- Endereço: http://127.0.0.1:8765 (atalho "Diário de Atividades" no menu de aplicativos)
- Dados: `data/atividades.db` (SQLite). **Backup = copiar esse arquivo.**
- Sem dependências: só Python 3 da própria máquina.

## Serviço (inicia sozinho no login)

```bash
systemctl --user status diario-atividades     # ver status
systemctl --user restart diario-atividades    # reiniciar (após alterar o código)
systemctl --user disable --now diario-atividades  # desligar
journalctl --user -u diario-atividades        # logs
```
