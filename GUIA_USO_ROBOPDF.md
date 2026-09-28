# RoboPDF - Passo a passo de uso

## Visao geral

Ao abrir o RoboPDF, voce vera 3 abas principais:

```text
[ Dividir PDF ]   [ Juntar PDFs ]   [ Extrair Paginas ]
```

Use cada aba conforme a tarefa que deseja fazer.

---

## 1. Dividir PDF

Use esta parte para separar um PDF grande em varios arquivos menores.

### Caminho na tela

```text
Aba: Dividir PDF
1. Selecione o PDF
2. Escolha como dividir
3. Gerar PDFs divididos
```

### Passo a passo

1. Clique em `Procurar PDF`.
2. Escolha o arquivo PDF no computador.
3. Confira a mensagem com o total de paginas carregadas.
4. Escolha um modo de divisao:

```text
Automatico
Divide o PDF em 2 a 9 partes.

Personalizado
Voce informa quantas paginas tera cada bloco.
```

### Modo automatico

```text
PDF carregado
   |
   v
Escolha: Dividir em 2, 3, 4... ate 9 partes
   |
   v
Clique em GERAR PDFs DIVIDIDOS
   |
   v
Escolha a pasta onde salvar
```

Exemplo:

```text
PDF com 100 paginas
Dividir em 4 partes
Resultado aproximado:
Parte_1_25pags.pdf
Parte_2_25pags.pdf
Parte_3_25pags.pdf
Parte_4_25pags.pdf
```

### Modo personalizado

```text
PDF carregado
   |
   v
Digite a quantidade de paginas do bloco
   |
   v
Clique em Adicionar Bloco
   |
   v
Repita ate alocar as paginas desejadas
   |
   v
Clique em GERAR PDFs DIVIDIDOS
```

Exemplo:

```text
PDF com 50 paginas

Bloco 1: 10 paginas
Bloco 2: 15 paginas
Bloco 3: 25 paginas

Resultado:
Parte_1_10pags.pdf
Parte_2_15pags.pdf
Parte_3_25pags.pdf
```

Se errar um bloco, use `Desfazer Ultimo Bloco`.

---

## 2. Juntar PDFs

Use esta parte para unir dois ou mais PDFs em um unico arquivo.

### Caminho na tela

```text
Aba: Juntar PDFs
1. Adicione os PDFs
2. Ordene a lista
3. Junte os arquivos
```

### Passo a passo

1. Clique em `Selecionar PDFs`.
2. Escolha dois ou mais arquivos PDF.
3. Confira a ordem na lista.
4. Se precisar, selecione um arquivo e use:

```text
Subir    -> move o PDF para cima
Descer   -> move o PDF para baixo
Remover  -> tira o PDF da lista
```

5. Clique em `JUNTAR PDFs AGORA`.
6. Escolha o nome e o local do arquivo final.

### Exemplo visual

```text
Lista na tela:

1. capa.pdf
2. contrato.pdf
3. anexos.pdf

Resultado:
pdf_unificado.pdf
```

A ordem da lista sera a ordem das paginas no PDF final.

---

## 3. Extrair Paginas

Use esta parte para criar um novo PDF contendo apenas algumas paginas de um PDF original.

### Caminho na tela

```text
Aba: Extrair Paginas
1. Selecione o PDF de origem
2. Informe as paginas
3. Gere o novo PDF
```

### Passo a passo

1. Clique em `Procurar PDF`.
2. Escolha o PDF de origem.
3. Confira o total de paginas carregadas.
4. No campo de paginas, digite as paginas desejadas.
5. Clique em `EXTRAIR E GERAR NOVO PDF`.
6. Escolha onde salvar o novo arquivo.

### Formatos aceitos

```text
1
Extrai apenas a pagina 1.

1, 5, 9
Extrai as paginas 1, 5 e 9.

10-15
Extrai da pagina 10 ate a 15.

1, 5, 10-15, 20
Mistura paginas soltas e intervalos.
```

### Exemplo visual

```text
PDF original: 30 paginas
Campo digitado: 1, 3, 10-12

Novo PDF tera:
pagina 1
pagina 3
pagina 10
pagina 11
pagina 12
```

---

## Erros comuns

```text
"Selecione um PDF primeiro"
Voce tentou executar uma acao sem carregar um PDF.

"Adicione pelo menos 2 PDFs"
Para juntar, e preciso selecionar no minimo dois arquivos.

"Formato invalido"
Na extracao, use apenas numeros, virgulas e tracos.

"Nenhuma pagina valida"
As paginas digitadas nao existem no PDF carregado.
```

---

## Resumo rapido

```text
Quero separar um PDF      -> Dividir PDF
Quero unir varios PDFs    -> Juntar PDFs
Quero pegar so paginas    -> Extrair Paginas
```
