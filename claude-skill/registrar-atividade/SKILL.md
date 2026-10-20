---
name: registrar-atividade
description: Registra no Diário de Atividades (app local em 127.0.0.1:8765) o que foi feito no dia. Use quando o usuário pedir para registrar, anotar ou lançar atividades, "o que eu fiz hoje", ou ao fim de uma tarefa concluída.
---

# Registrar atividade no Diário

O Diário de Atividades guarda o que foi feito no dia. No fim do mês, ele gera o relatório .docx para o cliente.
Cada lançamento vira matéria-prima desse relatório. Por isso, a descrição precisa se sustentar sozinha.

## 1. Juntar o que foi feito

Use, nesta ordem, o que estiver disponível:
- o que o usuário contou na mensagem;
- o trabalho feito nesta conversa;
- os commits do dia no repositório atual: `git log --since=midnight --author="$(git config user.email)" --stat`.

Um lançamento por frente de trabalho, não por commit. Juntar commits do mesmo assunto.

## 2. Montar cada lançamento

Campos (JSON):
- `data`: AAAA-MM-DD. É hoje, salvo se o usuário disser outro dia.
- `projeto`: nome do repositório ou cliente. Reaproveitar os nomes que já existem (passo 3).
- `categoria`: uma de `Desenvolvimento`, `Correção de bug`, `Testes`, `Refatoração`, `Análise / Investigação`, `Documentação`, `Infraestrutura / DevOps`, `Reunião`.
- `ticket`: id do card/história (ex.: `#359195`), se houver. Procurar no nome da branch e nas mensagens de commit.
- `descricao`: ver o estilo abaixo.
- `entregavel`: `true` se algo foi concluído e entregue (PR aberto/mergeado, história encerrada, relatório entregue, versão publicada). Senão, `false`.
- `horas`: só se o usuário informar.

Estilo da `descricao` (português do Brasil, tom de relatório para o cliente):
- Começar pelo verbo no particípio: "Corrigido…", "Implementado…", "Validado…", "Analisado…".
- Dizer o problema, o que mudou e o resultado, com números quando houver (testes, cobertura, medições).
- Explicar o termo técnico na primeira vez que aparecer. Nada de nome de função ou caminho de arquivo, a menos que seja essencial.
- Nunca incluir segredos, URLs assinadas, dados de pacientes ou conteúdo de documentos do acervo.
- Entre 2 e 8 frases. Pendências e decisões ficam no fim ("ficou recomendado como item separado").

## 3. Consultar os nomes já usados

```bash
curl -s http://127.0.0.1:8765/api/sugestoes
```

## 4. Mostrar e enviar

Mostre ao usuário a lista curta (data · projeto · categoria · ticket · 1ª frase · entrega?) e envie em seguida, a menos que ele peça ajuste:

```bash
curl -s -X POST http://127.0.0.1:8765/api/atividades \
  -H 'Content-Type: application/json' \
  --data-binary @/tmp/atividades.json
```

O arquivo pode ter um objeto ou uma lista de objetos. Grave o JSON num arquivo em vez de colocá-lo inline, por causa das aspas na descrição.

Se o app não responder (serviço parado), grave direto no banco:

```bash
python3 - <<'PY'
import json, sqlite3, os
db = os.path.expanduser("~/Documents/indra/diario-atividades/data/atividades.db")
itens = json.load(open("/tmp/atividades.json"))
itens = itens if isinstance(itens, list) else [itens]
with sqlite3.connect(db) as c:
    for i in itens:
        c.execute("INSERT INTO atividades (data, projeto, categoria, ticket, descricao, horas, entregavel) "
                  "VALUES (?,?,?,?,?,?,?)", (i["data"], i.get("projeto",""), i.get("categoria",""),
                  i.get("ticket",""), i["descricao"], i.get("horas",0), int(bool(i.get("entregavel")))))
print("ok")
PY
```
(Se a coluna `entregavel` não existir, reinicie o app uma vez: `systemctl --user restart diario-atividades`.)

## 5. Confirmar

Responda em uma linha: quantos lançamentos foram feitos e o link http://127.0.0.1:8765/#hoje.
