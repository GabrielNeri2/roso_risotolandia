# Robô de reservas — Risotolândia (Solvis)

Robô simples que faz as reservas de refeição no formulário
[https://web.solvis.net.br/s/reserva_clear_neo](https://web.solvis.net.br/s/reserva_clear_neo)
para todas as pessoas listadas na planilha `robo_risotolandia.xlsx`.

Ele usa **apenas o PowerShell nativo do Windows** — não instala nada, não baixa
nada, não abre navegador e não injeta nada em outros programas. Por isso é uma
abordagem de baixíssimo risco para antivírus/EDR como o CrowdStrike: são só
requisições HTTPS normais feitas por uma ferramenta assinada da Microsoft,
usando o "modo de compatibilidade" que o próprio site da Solvis oferece.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `robo_risotolandia.cmd` | **Duplo clique aqui para rodar de verdade** (envia as reservas) |
| `testar_sem_enviar.cmd` | Modo teste: percorre o formulário inteiro mas **não** clica em "Finalizar" |
| `robo_risotolandia.ps1` | O script em si (chamado pelos dois `.cmd` acima) |
| `robo_risotolandia.xlsx` | Planilha com as opções e os nomes |
| `log_robo_AAAA-MM-DD.txt` | Registro do que aconteceu em cada execução (criado automaticamente) |

## Como usar

1. Confira/edite a planilha `robo_risotolandia.xlsx` (uma linha por pessoa).
2. Rode `testar_sem_enviar.cmd` uma vez para conferir que está tudo casando
   com o site (nenhuma reserva é criada nesse modo).
3. Rode `robo_risotolandia.cmd`. O robô faz o formulário completo uma vez por
   nome: unidade → prato → refeição → horário → (dia, se for sexta) → nome →
   **Finalizar**.

## Formato da planilha

Colunas na primeira linha (a ordem não importa, e o robô reconhece o cabeçalho
mesmo com pequenas variações):

| empresa | prato | refeicao | horario | Se for na sexta | Nome |
|---|---|---|---|---|---|
| ClearCorrect | Selezione | Almoco | 11:30 | Segunda-feira | Andre Aliot |

- Linhas **sem nome são ignoradas** — para tirar alguém da reserva do dia,
  basta apagar o nome da linha.
- A comparação com as opções do site tolera acento, maiúscula/minúscula,
  espaço e pequenas diferenças de grafia (ex.: "Selezione" na planilha casa
  com "Selezioni" no site, "Almoco" casa com "Almoço").

## Regra da sexta-feira

Na sexta-feira o formulário mostra uma pergunta extra — "Sua reserva é para:"
com as opções **Sábado (amanhã)** e **Segunda-feira (da próxima semana)**.
O robô seleciona o que estiver na coluna **"Se for na sexta"** da planilha
(por padrão, `Segunda-feira`). Se a coluna estiver vazia, ele escolhe
Segunda-feira mesmo assim. Nos outros dias essa pergunta não aparece e o robô
simplesmente segue em frente.

## Agendar para rodar sozinho (opcional)

No **Agendador de Tarefas** do Windows, crie uma tarefa de segunda a sexta no
horário desejado com a ação:

```
Programa:   powershell.exe
Argumentos: -NoProfile -ExecutionPolicy Bypass -File "C:\caminho\da\pasta\robo_risotolandia.ps1"
Iniciar em: C:\caminho\da\pasta
```

## Problemas comuns

- **"Não foi possível ler a planilha"** — o robô lê o `.xlsx` pelo Excel
  instalado na máquina. Se o Excel não estiver instalado, salve a planilha
  também como `robo_risotolandia.csv` (CSV) na mesma pasta que o robô a usa
  automaticamente.
- **Alguma pessoa falhou** — veja o arquivo `log_robo_AAAA-MM-DD.txt`: ele
  mostra em qual pergunta parou e quais opções o site ofereceu. Normalmente é
  um valor na planilha que não existe mais no site (ex.: horário esgotado).
- **Rodou duas vezes no mesmo dia** — o robô não detecta reservas repetidas;
  cada execução envia tudo de novo. Rode uma vez por dia.
