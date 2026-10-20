# Diário de Atividades

App local para anotar o que foi feito no dia e gerar o relatório mensal para o cliente.

- Endereço: http://127.0.0.1:8765 (atalho "Diário de Atividades" no menu de aplicativos)
- Dados: `data/atividades.db` (SQLite). **Backup = copiar esse arquivo.**
- Sem dependências: só Python 3 da própria máquina.

## Relatório .docx com IA

1. Aba **Modelos**: envie o padrão da empresa (.docx/.dotx) e um ou mais relatórios de referência (.docx/.dotx/.pdf).
   Ficam em `modelos/padrao/` e `modelos/referencia/` e não vão para o Git.
2. Copie `.env.example` para `.env` e preencha `OPENAI_API_KEY` (e, se quiser, `OPENAI_MODEL`). Reinicie o app.
3. Aba **Relatório do mês** → **Gerar texto com IA**: a OpenAI resume as atividades no estilo das referências.
4. Revise o texto e clique em **Baixar .docx**: o documento é montado sobre o padrão (capa, índice, estilos, rodapé).
   Ao abrir no Word, confirme a atualização dos campos para o índice ganhar os números de página.

Dados da capa (título, autor, área, versão) ficam em `data/config.json`; rascunhos em `data/rascunhos/`.

## Serviço (inicia sozinho no login)

```bash
systemctl --user status diario-atividades     # ver status
systemctl --user restart diario-atividades    # reiniciar (após alterar o código)
systemctl --user disable --now diario-atividades  # desligar
journalctl --user -u diario-atividades        # logs
```
