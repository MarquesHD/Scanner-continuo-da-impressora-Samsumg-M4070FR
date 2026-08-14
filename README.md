# Scanner Continuo Samsung M4070FR → PDF

App simples e com bom design para digitalizar documentos direto da
impressora Samsung M4070FR e salvar sempre em PDF, com nome de
arquivo editável e modo de digitalização contínua (várias páginas no
mesmo PDF).

Feito em Python (interface com `customtkinter`, comunicação com o
scanner via WIA — Windows Image Acquisition).

<img width="819" height="747" alt="image" src="https://github.com/user-attachments/assets/fba2bc0d-179c-4679-957d-4f693eb78a96" />

---

## 1. Pré-requisitos

- Windows 10/11
- Python 3.9+ instalado ([python.org](https://www.python.org/downloads/))
  - No instalador, marque a opção **"Add Python to PATH"**
- Driver da Samsung M4070FR instalado, com a função de scanner
  reconhecida pelo Windows (teste abrindo o "Fax e Scanner do
  Windows" ou o app "Digitalizar" da Microsoft Store — se a
  impressora aparecer lá, o app vai funcionar)

## 2. Como rodar (modo rápido, sem gerar .exe)

Abra o Prompt de Comando (cmd) dentro da pasta do projeto e rode:

```bat
pip install -r requirements.txt
python main.py
```

A janela do app deve abrir. Pronto, já pode digitalizar.

## 3. Como gerar um .exe instalável para usar sem o Python aberto

Dentro da pasta do projeto, dê **duplo clique** em `build_exe.bat`
(ou rode pelo cmd). Ele vai:

1. Instalar as dependências
2. Instalar o PyInstaller
3. Gerar `dist\ScannerSamsung.exe`

Esse `.exe` pode ser copiado para a Área de Trabalho ou fixado na
barra de tarefas — não precisa mais do Python instalado na máquina
que for usar (só na máquina que gerou o .exe).

> Obs: como não tenho como testar em uma máquina Windows real com a
> impressora conectada, é bem possível que algum ajuste pequeno seja
> necessário na primeira tentativa (driver, nome exato do
> dispositivo, etc.). Se der erro, me manda a mensagem que eu
> ajusto o código.

## 4. Como usar o app

1. Digite o **nome do arquivo** (vira `nome.pdf` automaticamente)
2. Escolha a **pasta de destino** (padrão: `Documentos\Digitalizados`)
3. Deixe **"Digitalização contínua"** ligada se for escanear várias
   folhas para o mesmo PDF
4. Clique em **"Digitalizar página"** — repita para cada folha
5. Errou uma página? Clique no ícone 🗑️ ao lado dela na lista para
   removê-la só ela, sem descartar as outras
6. Quando terminar, clique em **"Finalizar e salvar PDF"**
   - Isso junta todas as páginas digitalizadas (na ordem da lista)
     em um único PDF
7. Para descartar tudo e recomeçar, use **"Cancelar sessão"**
8. O botão 🌙/☀️ no canto superior direito alterna entre tema claro
   e escuro

## 5. Problemas comuns

- **"Nenhum scanner encontrado"**: verifique se a impressora está
  ligada, conectada (USB ou rede) e se o driver de scanner (não só
  de impressão) está instalado. Teste primeiro no app nativo
  "Digitalizar" do Windows, e instale o driver Universal.

  https://support.hp.com/us-en/drivers/samsung-proxpress-sl-m4070-laser-multifunction-printer-series/model/16462932
<img width="1156" height="306" alt="image" src="https://github.com/user-attachments/assets/c4868eb8-614e-41d5-8627-ba933ff53580" />

  Após a instalação do Driver, é necessário incluir a impressora como Scanner.
<img width="702" height="459" alt="image" src="https://github.com/user-attachments/assets/9e3ab7ee-d8f3-4772-a090-42c8b8768c71" />
  
- **Erro ao gerar o .exe**: rode `pip install pywin32` de novo e, se
  necessário, o script `python Scripts\pywin32_postinstall.py -install`
  que vem com o pywin32 (normalmente em
  `%APPDATA%\..\Local\Programs\Python\PythonXX\Scripts`).
- **Imagem digitalizada com qualidade ruim**: o app usa as
  configurações padrão do driver da impressora. Se quiser, dá pra eu
  adicionar um seletor de resolução (DPI) e cor/preto-e-branco.

## 6. Estrutura do projeto

```
scanner_app/
├── main.py            # app principal (interface + lógica)
├── requirements.txt   # dependências Python
├── build_exe.bat       # script para gerar o .exe
└── README.md           # este arquivo
```
