# MAGRE BANK · Controle Financeiro

Dashboard financeiro pessoal em dark mode premium, em um único `index.html` (Tailwind + Lucide + Montserrat via CDN), com sincronização opcional no Google Planilhas.

Sem planilha conectada, tudo funciona em **modo local** (dados salvos no navegador).

---

## 1. Conectar ao Google Planilhas

1. Crie uma planilha nova no Google Planilhas (ex.: "Magre Bank").
2. Vá em **Extensões → Apps Script**.
3. Apague o conteúdo do editor e cole o arquivo [`apps-script/Code.gs`](apps-script/Code.gs).
4. Na linha `const CHAVE = '...'`, troque pela sua senha secreta.
5. No topo do editor, escolha a função **`configurar`** e clique em **Executar**. Autorize o acesso quando pedir. As abas `Lancamentos`, `Cartoes`, `Alvos` e `Config` serão criadas.
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

## 2. Publicar no GitHub Pages

1. Crie um repositório no GitHub e envie estes arquivos.
2. No repositório: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)`** → Save.
3. Em 1 ou 2 minutos o site fica disponível em `https://SEU-USUARIO.github.io/NOME-DO-REPO/`.

## Estrutura

```
index.html          → o site completo
apps-script/Code.gs → API da planilha (cole no Apps Script)
```
