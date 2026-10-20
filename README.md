# Diário de Atividades

App local para anotar o que foi feito no dia e gerar o relatório mensal para o cliente.

- Endereço: http://127.0.0.1:8765 (atalho "Diário de Atividades" no menu de aplicativos)
- Dados: `data/atividades.db` (SQLite). **Backup = copiar esse arquivo.**
- Sem dependências: só Python 3 da própria máquina.

## Relatório .docx com IA

1. **Configurações → Modelos de relatório**: crie quantos modelos quiser (ex.: um por cliente). Cada modelo tem
   o próprio template (.docx/.dotx), até 10 relatórios de referência (.docx/.dotx/.pdf) e os dados da capa.
   Associe projetos a um modelo para ele ser escolhido sozinho; o modelo *principal* vale para o resto.
   Os arquivos ficam em `modelos/perfis/<modelo>/` e não vão para o Git.
2. **Configurações → Inteligência artificial**: escolha o provedor (OpenAI, Anthropic ou qualquer API compatível
   com OpenAI — Azure, OpenRouter, Groq, Ollama…), o modelo e a chave. Também dá para usar o `.env` (veja `.env.example`).
3. **Relatório do mês**: escolha o mês, o projeto e o modelo → **Gerar com IA** → revise → **Baixar .docx**.
   Ao abrir no Word, confirme a atualização dos campos para o índice ganhar os números de página.

### Template próprio

Qualquer .docx/.dotx serve. Marcadores preenchidos na capa, cabeçalho e rodapé:
`{{titulo}} {{subtitulo}} {{autor}} {{area}} {{versao}} {{mes}} {{projeto}} {{data}}`.
O texto entra no parágrafo `{{conteudo}}`; sem ele, substitui o conteúdo de exemplo a partir do primeiro
Título 1 depois do sumário; sem isso, vai para o fim. Títulos e listas usam os estilos Título 1–4 e
Lista com marcadores do template.

Tema claro/escuro e cor de destaque: **Configurações → Aparência** (ou o botão da lua na barra lateral).

## Serviço (inicia sozinho no login)

```bash
systemctl --user status diario-atividades     # ver status
systemctl --user restart diario-atividades    # reiniciar (após alterar o código)
systemctl --user disable --now diario-atividades  # desligar
journalctl --user -u diario-atividades        # logs
```

## Registro pelo Claude Code

A skill `claude-skill/registrar-atividade` transforma "o que eu fiz hoje" (ou os commits do dia) em lançamentos
e envia para `POST /api/atividades` (aceita um objeto ou uma lista). Para instalar:

```bash
mkdir -p ~/.claude/skills && ln -s "$PWD/claude-skill/registrar-atividade" ~/.claude/skills/registrar-atividade
```

Depois, no Claude Code: "registra o que eu fiz hoje".
Marque como **entrega** o que foi concluído: a IA destaca essas atividades no relatório do mês.
