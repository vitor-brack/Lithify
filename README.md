# Lithify

Compressor de imagens para laudos técnicos em Word.

Reduz fotos de campo em alta resolução para tamanhos práticos sem perda perceptível, evitando que o Word trave ou que o documento final fique grande demais para enviar por e-mail.

Tem dois modos:

- **Comprimir Imagens** *(recomendado)* — comprime as fotos da pasta antes de você inserir no Word.
- **Comprimir Word** *(válvula de escape)* — recebe um `.doc` ou `.docx` pronto e comprime as imagens já embutidas.

---

## Para o usuário final

### Instalação

1. Baixe o arquivo `Lithify.exe`.
2. Coloque-o numa pasta da sua escolha (ex.: `C:\Users\Seu Nome\Programas\Lithify\`).
3. Dê um duplo-clique para abrir.

Não precisa instalar Python nem nada extra. O arquivo é único e portátil.

> Na primeira vez, o Windows pode mostrar um aviso de segurança ("Windows Defender SmartScreen"). Clique em **Mais informações → Executar assim mesmo**.

### Modo 1 — Comprimir Imagens (recomendado)

Use **antes** de inserir as fotos no Word.

1. Abra o Lithify.
2. Na aba **Comprimir Imagens** (já vem selecionada), arraste as fotos ou a pasta inteira para a área grande no topo. Você também pode clicar em **Selecionar arquivos** ou **Selecionar pasta**.
3. Escolha o preset:
   - **Tela / PDF (recomendado)** — para laudos que serão lidos no computador ou exportados em PDF.
   - **Impressão** — quando o laudo será impresso em papel.
   - **Personalizado** — controle fino da resolução e qualidade.
4. (Opcional) Escolha uma pasta de saída específica. Se deixar vazio, as fotos comprimidas vão para uma subpasta `comprimidas/` ao lado dos originais.
5. Clique em **Comprimir**.

As fotos originais **não são alteradas** (a menos que você marque a caixa "Sobrescrever os arquivos originais", que pede confirmação).

Depois, é só inserir as fotos da pasta `comprimidas/` no seu documento Word normalmente.

### Modo 2 — Comprimir Word (válvula de escape)

Use quando você já montou o documento com as fotos sem ter comprimido antes. Aceita os dois formatos: `.doc` (formato antigo) e `.docx` (formato moderno).

> ⚠️ **Antes de usar, faça uma cópia de segurança do seu documento original.** O Lithify gera um arquivo novo numa pasta separada, mas é boa prática ter backup de qualquer documento importante.

1. Vá para a aba **Comprimir Word**.
2. Arraste o arquivo `.doc` ou `.docx` para a área de drop, ou clique para selecionar.
3. Escolha o preset (mesmas opções do Modo 1).
4. Clique em **Comprimir Word**.

O resultado vai para uma subpasta `comprimidos/` ao lado do arquivo original, com o **mesmo nome e mesmo formato** do original (`.doc` continua `.doc`, `.docx` continua `.docx`).

#### Sobre arquivos `.doc`

Para comprimir `.doc`, o Lithify usa o **Microsoft Word instalado no seu computador** como conversor invisível. Fluxo:

1. Word abre o `.doc` em segundo plano e salva como `.docx` temporário.
2. Lithify comprime as imagens.
3. Word converte de volta para `.doc` no destino final.

O processo adiciona alguns segundos a mais (porque o Word precisa abrir e fechar). É normal.

Se você não tiver Word instalado, o Lithify avisa com uma mensagem clara e sugere o workflow alternativo: abrir o `.doc` no Word, **Arquivo → Salvar como → Documento do Word (`.docx`)**, e usar o Lithify com o `.docx` resultante.

### Dicas

- **Quanto reduzo?** No preset padrão, fotos de 4K geralmente caem de ~8 MB para ~300-500 KB cada — sem que dê para perceber a diferença a olho nu em laudo impresso ou PDF.
- **PNG com fundo transparente?** Por padrão, o Lithify converte para JPEG com fundo branco (mais compacto). Se a transparência for importante, marque a caixa "Preservar transparência".
- **Imagens já otimizadas** (JPEG pequeno e dentro do limite de tamanho) são copiadas sem reencodar — você não perde qualidade tentando comprimir o que já está bom.
- **Imagens vetoriais** (EMF, WMF, SVG) dentro do `.docx` são preservadas intactas. Não há perda nessas.
- **"Bitmaps embutidos" (OLE Paint)**: às vezes uma imagem é colada no Word com Ctrl+V de fonte errada e fica armazenada como objeto OLE Paint (Bitmap Image). Isso explode o tamanho do arquivo (cada foto vira BMP cru de ~26 MB). O Lithify detecta e comprime esses casos também — o double-click ainda abre Paint normalmente, só com a imagem em resolução menor.

---

## Para o desenvolvedor

### Setup

```powershell
git clone <repo>
cd Lithify
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Requer Python 3.10+.

### Estrutura

```
Lithify/
├── core/
│   ├── exceptions.py
│   ├── image_compressor.py    # núcleo: comprimir_imagem + comprimir_lote
│   ├── ole_processor.py       # rewrite de OLE PBrush (Bitmap Image)
│   ├── docx_processor.py      # comprimir_docx: word/media + word/embeddings
│   ├── doc_converter.py       # automação Word/COM (.doc ↔ .docx)
│   └── document_processor.py  # orquestrador: detecta formato e despacha
├── gui/
│   └── main_window.py         # MainWindow + ImageTab + DocxTab
├── tests/
│   ├── test_image_compressor.py
│   ├── test_docx_processor.py
│   ├── test_document_processor.py
│   └── test_ole_processor.py
├── main.py                    # ponto de entrada + logging
├── requirements.txt
├── pytest.ini
└── README.md
```

`core/` é puro (sem GUI, sem `print`, sem side-effects globais), com type hints e callback de progresso opcional. A lógica de compressão de imagem mora apenas em `core/image_compressor.py` — `docx_processor.py` reutiliza, não reimplementa.

A camada de Word/COM está isolada em `core/doc_converter.py`. O orquestrador `core/document_processor.py` é o único ponto que conhece os dois formatos: detecta `.doc` vs `.docx` por magic bytes e dispatcha. A GUI fala apenas com o orquestrador.

#### Compressão de OLE Bitmap Images (`core/ole_processor.py`)

Quando uma imagem é colada no Word como "Bitmap Image" (objeto OLE Paint), ela vai parar em `word/embeddings/oleObject*.bin` em vez de `word/media/`. Cada `.bin` é um OLE compound file com CLSID PBrush (`0003000A-...`) contendo dois streams pesados:

- `\x01Ole10Native`: BMP cru de até 26 MB para uma foto 4K
- `\x03PRINT`: cópia metafile do mesmo bitmap, mais 26 MB

O `ole_processor` reescreve o compound usando a Win32 Structured Storage API (`pythoncom.StgCreateDocfile`):

1. Extrai o BMP de `Ole10Native`, redimensiona com Pillow.
2. Reconstrói o compound preservando CLSID, `\x01CompObj`, `\x01Ole`, etc.
3. Descarta `\x03PRINT` (o Word regenera automaticamente ao salvar como `.doc`).

O double-click na imagem ainda abre o Paint normalmente, só com bitmap menor. Em um laudo real do projeto: 689 MB → 73 MB (89% de redução).

### Rodar os testes

```powershell
python -m pytest tests/ -v
```

Atualmente: 38 testes (compressor de imagem + processor de .docx + orquestrador .doc/.docx + ole_processor). Os fixtures de documento são gerados programaticamente em memória via `zipfile`, Pillow e `pythoncom.StgCreateDocfile` — não há binários no repositório.

Testes que dependem de Win32 Structured Storage (`test_ole_processor.py`) ou de Microsoft Word (`test_integracao_doc_round_trip_produz_doc_valido`) são pulados automaticamente em ambientes sem essas dependências (`pytest.mark.skipif`).

### Logging

O `main.py` configura um `RotatingFileHandler` em:

- Windows: `%APPDATA%\Lithify\lithify.log`
- Outros: `~/.lithify/lithify.log`

Para logging detalhado no console, rode com `--debug`:

```powershell
python main.py --debug
```

### Empacotamento (PyInstaller)

Comando testado em Windows 11 + Python 3.13:

```powershell
pyinstaller --onefile --windowed --name Lithify --collect-all tkinterdnd2 --collect-submodules win32com --collect-submodules pythoncom --hidden-import olefile --noconfirm main.py
```

O resultado fica em `dist/Lithify.exe` (~26 MB). Flags importantes:

- `--collect-all tkinterdnd2`: a `tkinterdnd2` carrega DLLs nativas (`tkdnd`) em runtime; sem isso o drag-and-drop quebra na primeira tentativa.
- `--collect-submodules win32com --collect-submodules pythoncom`: garantem que pywin32 funcione no executável empacotado (necessário para abrir/converter `.doc` via Word **e** para reescrever OLE compound files via Structured Storage).
- `--hidden-import olefile`: PyInstaller pode não detectar `olefile` automaticamente; este flag garante a inclusão.

Se quiser um `.exe` com ícone próprio, adicione `--icon caminho/para/icone.ico`.

Para limpar artefatos do build:

```powershell
Remove-Item -Recurse -Force build, dist, Lithify.spec
```

### Como contribuir

1. Crie um branch a partir de `main`.
2. Adicione/atualize testes em `tests/` para cobrir o que mudou.
3. Confira `python -m pytest tests/`.
4. Abra o PR descrevendo a mudança e o motivo.

A regra mais importante: **toda lógica de compressão fica em `core/`** e é testável sem GUI. A GUI é só uma camada fina sobre o `core/`.
