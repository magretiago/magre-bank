# MAGRE BANK · Controle Financeiro

Dashboard financeiro pessoal em dark mode premium, em um único `index.html` (Tailwind + Lucide + Montserrat via CDN), com sincronização opcional no Google Planilhas.

Sem planilha conectada, tudo funciona em **modo local** (dados salvos no navegador).

---

## 1. Conectar ao Google Planilhas

1. Crie uma planilha nova no Google Planilhas (ex.: "Magre Bank").
2. Vá em **Extensões → Apps Script**.
3. Apague o conteúdo do editor e cole o arquivo [`apps-script/Code.gs`](apps-script/Code.gs).
4. Na linha `const CHAVE = '...'`, troque pela sua senha secreta.
5. No topo do editor, escolha a função **`configurar`** e clique em **Executar**. Autorize o acesso quando pedir. As abas `Lancamentos`, `ContasBancarias`, `Cartoes`, `Faturas`, `Alvos`, `Contas`, `Categorias` e `Config` serão criadas, e o lembrete diário por e-mail (8h) é ativado.
6. Clique em **Implantar → Nova implantação**:
   - Tipo: **App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
7. Copie a **URL do App da Web** (termina em `/exec`).
8. No site, clique na **engrenagem** ao lado do indicador de sincronização, cole a URL e a chave, e clique em **Testar e conectar**.

Na primeira conexão, se a planilha estiver vazia, o site envia os dados atuais para ela. Depois disso, cada lançamento é gravado na planilha, e o botão de sincronização busca o que estiver lá.

> A URL e a chave ficam salvas apenas no navegador de quem conectou. Elas **não** vão para o código no GitHub.
>
> Se você alterar o `Code.gs` depois, faça **Implantar → Gerenciar implantações → Editar → Nova versão** para a URL passar a usar o código novo.

## Contas a pagar e lembretes

- Cadastre contas recorrentes (aluguel, água, energia, internet…) com o dia do vencimento. O valor é opcional, para contas de consumo.
- O site mostra um alerta no topo e um contador no sino quando há contas vencidas ou vencendo nos próximos dias.
- Ao clicar em **Pagar**, informe o valor, por onde pagou (uma conta bancária ou um cartão) e a forma (Pix, boleto, dinheiro…). Você pode atualizar o valor padrão da conta. O pagamento vai direto para o extrato e a conta passa para o próximo mês.
- Remover o lançamento do pagamento no extrato reabre a conta.
- Todo dia às 8h o script envia um e-mail se houver contas vencidas ou vencendo. Para testar na hora, rode a função **`testarLembrete`** no Apps Script.
- Na aba `Config` você pode mudar o `email` de destino e os `diasAviso` (antecedência do aviso).
- Confira o fuso do projeto em **Apps Script → Configurações do projeto → Fuso horário** (deixe *São Paulo*).

## Contas bancárias, transferências e categorias

- **Contas bancárias:** cadastre cada conta por onde recebe e paga (corrente, digital, poupança ou carteira de dinheiro), com o saldo inicial. O saldo atual de cada uma é calculado pelos lançamentos, e o card "Saldo Disponível" mostra a soma.
- **Forma de pagamento:** receitas e despesas registram como o dinheiro entrou ou saiu (Pix, débito, dinheiro, boleto, transferência…). O extrato mostra, por exemplo, "Itaú · Pix", e dá para filtrar por conta e por forma.
- **Transferência:** move dinheiro entre suas contas (um saque para a carteira, um Pix entre contas suas) sem contar como receita nem despesa.
- **Categorias:** no menu **Categorias** (ou na engrenagem ao lado do campo Categoria) você cria, renomeia e remove categorias de cada tipo. Ao renomear, os lançamentos antigos são atualizados.

## Cobrança automática: assinaturas no cartão e débito automático

Ao cadastrar (ou editar, no lápis) uma conta a pagar, escolha **como ela é cobrada**:

- **Pagamento manual** (boleto, Pix): alerta, lembrete por e-mail e botão "Pagar".
- **No cartão de crédito**: no dia da cobrança, o valor é lançado sozinho como gasto no cartão escolhido e entra na fatura certa. Não gera alerta de pagamento (a fatura já tem o seu).
- **Débito automático**: no dia do débito, o valor é lançado sozinho na conta bancária escolhida. Nos dias anteriores, o site avisa "Débito em X dias" e, se o saldo da conta não for suficiente, mostra o alerta em vermelho (também no e-mail).
- **Valor variável** (energia, água): no dia, a conta pede para **confirmar o valor** cobrado ou debitado. O site usa o último valor pago como estimativa.
- Os lançamentos automáticos acontecem quando o site é aberto, inclusive os de dias ou meses que ficaram para trás, com a data certa.
- Excluir do extrato uma cobrança automática não faz ela ser lançada de novo.
- O resumo **Recorrentes por mês** mostra quanto você gasta com contas fixas, separado por cartão, débito automático e pagamento manual.

## Faturas de cartão e parcelamento

- Cada cartão tem **dia de fechamento** e **dia de vencimento**. Compras até o fechamento entram na fatura daquele mês; depois disso, na fatura seguinte. O cartão mostra a **fatura aberta** (com quantos dias faltam para fechar), a **fatura fechada a pagar** e o **melhor dia de compra**.
- **Parcelamento:** no gasto no cartão, escolha de 1x a 24x. Cada parcela cai numa fatura, e o limite fica comprometido pelo valor total até as faturas serem pagas.
- **Pagar fatura:** pelo botão no cartão ou em Contas a Pagar. O valor sai da conta bancária escolhida e o limite é liberado. Pagamentos parciais são aceitos.
- As faturas fechadas aparecem em **Contas a Pagar**, no alerta do site e no lembrete por e-mail.
- No card **Despesas do Mês**, alterne entre **Por compra** (o que você comprou neste mês) e **Por fatura** (as faturas que vencem neste mês).
- No extrato, cada compra mostra a fatura em que caiu, e ao filtrar um cartão aparece o filtro por fatura.

### Atualizando o script (versões anteriores)

Sempre que o site pedir para atualizar o script (aviso vermelho no topo), siga os passos abaixo. Até lá, o site lê a planilha, mas não grava nada nela, e os lançamentos novos ficam guardados no navegador.

### Atualizando da versão 1 do script

Se a sua planilha usa o script antigo, o site mostra um aviso vermelho e **não grava nada** até você atualizar. Os lançamentos feitos nesse meio tempo ficam guardados no navegador e são enviados assim que o script novo estiver publicado.

1. Abra **Extensões → Apps Script** e substitua todo o código pelo novo `Code.gs`, mantendo a sua `CHAVE`.
2. Salve e rode **`configurar`** (cria as abas novas).
3. **Implantar → Gerenciar implantações →** lápis **(Editar) → Versão: Nova versão → Implantar.** A URL continua a mesma.
4. No site, clique no indicador de sincronização. Seus dados atuais passam a pertencer à conta "Conta Corrente Magre", com o saldo inicial que estava na aba `Config`. Depois é só renomear essa conta e cadastrar as outras.

## 2. Publicar no GitHub Pages

1. Crie um repositório no GitHub e envie estes arquivos.
2. No repositório: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)`** → Save.
3. Em 1 ou 2 minutos o site fica disponível em `https://SEU-USUARIO.github.io/NOME-DO-REPO/`.

## Estrutura

```
index.html          → o site completo
apps-script/Code.gs → API da planilha (cole no Apps Script)
```
