# RoboPDF - Download no Windows e geracao do .exe

## Diagnostico rapido

O problema de algumas pessoas nao conseguirem baixar ou executar no Windows nao parece ser erro do codigo do RoboPDF.

O motivo mais provavel e o Microsoft Defender SmartScreen:

```text
Arquivo novo
+ Sem assinatura digital
+ Poucos downloads ainda
+ Editor desconhecido
= Windows/Edge pode bloquear ou pedir confirmacao
```

Isso pode acontecer mesmo quando o arquivo e seguro.

## Por que alguns conseguem e outros nao?

Depende do computador de cada pessoa:

```text
Usuario comum
Pode aparecer a opcao de manter/executar mesmo assim.

Computador de empresa
A politica de seguranca pode impedir o usuario de continuar.

Windows 11 com Smart App Control
Pode bloquear arquivos sem assinatura de forma mais rigida.

Navegador diferente
Edge, Chrome e antivirus podem tratar o mesmo .exe de maneiras diferentes.
```

## Sobre o formulario da Microsoft

Preencher formulario ou enviar arquivo para analise da Microsoft nao libera o download imediatamente para todo mundo.

Esse processo serve para analise/revisao, mas a reputacao do SmartScreen normalmente depende de:

```text
assinatura digital
historico de downloads
downloads limpos por varias pessoas
reputacao do site de download
reputacao do certificado do publicador
```

Se voce gerar uma nova versao do .exe sem assinatura, a reputacao pode voltar praticamente a zero para aquele novo arquivo.

## Estado do arquivo do GitHub Actions

O workflow fica em:

```text
.github/workflows/build.yml
```

Ele foi ajustado para:

```text
1. Rodar no Windows
2. Instalar Python 3.11
3. Instalar dependencias do requirements.txt
4. Validar app.py
5. Gerar Robo_PDF.exe com PyInstaller
6. Incluir os arquivos do customtkinter
7. Conferir se dist/Robo_PDF.exe foi criado
8. Gerar checksum SHA256
9. Disponibilizar o .exe no GitHub Actions
10. Publicar no GitHub Releases quando voce criar uma tag v*
```

## Como gerar o executavel pelo GitHub

### Metodo automatico

```text
Commit no main
   |
   v
GitHub Actions roda sozinho
   |
   v
Baixe o artefato Robo_PDF.exe
```

Esse download de artefato e temporario.

### Metodo recomendado para distribuir

Use uma Release do GitHub.

Depois de enviar as alteracoes para o GitHub:

```bash
git tag v1.0.0
git push origin v1.0.0
```

Isso cria uma Release com:

```text
Robo_PDF.exe
checksums.txt
```

A Release e melhor para compartilhar com outras pessoas do que o artefato do Actions.

## O que orientar o usuario final

Use este texto curto:

```text
Baixe apenas pelo link oficial do GitHub Releases.
Se o Windows avisar que o app nao e conhecido, isso acontece porque o RoboPDF ainda nao tem assinatura digital/reputacao.
Confira se o arquivo veio do link oficial antes de continuar.
```

Em alguns computadores, principalmente de empresa, o usuario pode nao conseguir continuar sem ajuda do TI.

## Solucao definitiva

Para reduzir muito esses avisos, o ideal e:

```text
1. Assinar digitalmente o .exe com certificado Authenticode
2. Usar sempre o mesmo certificado/publicador
3. Distribuir por GitHub Releases ou Microsoft Store
4. Nao alterar o .exe depois de assinar
```

Sem assinatura digital, cada nova versao do Robo_PDF.exe pode voltar a ser tratada como arquivo desconhecido.

## Links oficiais

Microsoft SmartScreen para desenvolvedores:
https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation

FAQ do Microsoft Defender SmartScreen:
https://feedback.smartscreen.microsoft.com/smartscreenfaq.aspx

CustomTkinter com PyInstaller:
https://customtkinter.tomschimansky.com/documentation/packaging/

GitHub Actions Artifacts:
https://docs.github.com/en/actions/concepts/workflows-and-actions/workflow-artifacts

