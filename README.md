# Controle de Produção — base inicial

Site com login para controle de pedidos de produção, gerado a partir da planilha
`17_08_CONTROLE_PRODUCAO_V1.xlsx`. Todos os dados ficam num banco (SQLite) real —
nada é fixo/estático, tudo que aparece na tela vem do banco.

## O que já vem pronto

- **Login** (usuário `admin` / senha `admin123` — troque assim que possível, veja abaixo).
- **~1.146 pedidos da planilha já importados** automaticamente na primeira execução.
- **Painel** com indicadores (total de pedidos, pendentes, em tratativa, em
  andamento, finalizados, valor total) e filtros por cliente, vendedor, status e estação.
- **Tela "Novo pedido"**: só os campos que você pediu como base de inclusão —
  data do cliente, data de inclusão do pedido, cliente, CNPJ, cidade, estado,
  país, frete, vendedor, pedido de venda, descrição do produto, quantidade e
  custo unitário. O valor total (quantidade × custo) é calculado sozinho.
- **Tela de edição do pedido**: o restante das colunas (estação, status,
  prioridade, datas de produção, RNC, observações) fica ali, "parametrizado":
  - **Status** é automático — muda sozinho para `ANDAMENTO` quando você
    preenche "Início produção", e para `FINALIZADO` quando preenche "Término
    de inspeção" + "Liberação de faturamento". `EM TRATATIVA` é a única opção
    manual (fica travada até você finalizar o pedido).
  - **Lead times** (LT comercial, tempo de espera, LT produção, prazo total)
    são sempre calculados a partir das datas — nunca digitados.

## Como rodar

Pré-requisito: Python 3.10+ instalado.

```bash
cd producao_app
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Depois abra **http://localhost:5000** no navegador.

Na primeira execução o sistema cria o banco (`instance/pedidos.db`), o usuário
`admin` e importa a planilha `data/controle_producao_base.xlsx`. Isso só
acontece uma vez — nas próximas vezes ele só abre o banco já existente.

Se quiser recomeçar do zero (apagar tudo e importar a planilha de novo), basta
apagar o arquivo `instance/pedidos.db` e rodar `python app.py` outra vez.

## Estrutura do projeto

```
producao_app/
├── app.py               rotas (login, painel, novo pedido, editar pedido)
├── models.py             modelo dos dados + listas de opções dos dropdowns
├── seed.py                importação da planilha para o banco
├── extensions.py         configuração do banco/login
├── data/                  cópia da planilha original (usada só na 1ª execução)
├── templates/             páginas HTML
├── static/                CSS, JS e Bootstrap (tudo local, não depende de internet)
└── instance/              banco de dados SQLite (criado automaticamente)
```

## Pontos para você ajustar depois (é só pedir)

- **Listas dos dropdowns** (estações, status, prioridade, frete) estão todas
  no topo do `models.py` — fácil de adicionar/remover opções.
- **Regra do status automático** está em `models.py`, método
  `atualizar_status_automatico()` — hoje segue a lógica descrita acima, mas dá
  pra ajustar (por exemplo, criar uma regra pra "atrasado" comparando com
  "liberação prevista").
- **Múltiplos usuários/permissões**: hoje existe 1 usuário fixo (admin); dá
  para criar uma tela de cadastro de usuários.
- **Trocar a senha do admin**: por ora precisa ser feito direto no banco; se
  quiser já incluo uma tela de "trocar senha" na próxima versão.
- **Hospedagem**: este é um app Flask "de verdade" — hoje pensado para rodar
  na sua máquina, mas pode ser publicado num servidor (Render, PythonAnywhere,
  VPS, etc.) quando quiser deixá-lo acessível pela internet para a equipe.

## Segurança (importante antes de usar com dados reais / em rede)

- Troque a senha do usuário `admin`.
- Troque o valor de `SECRET_KEY` em `app.py` (ou defina a variável de
  ambiente `SECRET_KEY`) antes de usar fora do seu computador.
- O servidor embutido do Flask (`python app.py`) é só para uso local/teste —
  para publicar na internet, use um servidor de produção (ex: gunicorn).
