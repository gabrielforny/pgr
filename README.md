# RTF_to_DOCX
RPA que usa arquivo RTF como dados de entrada e Preenche um arquivo DOCX como saída.

# Instruções de Uso 

## REQUISITOS
 - 1 
```txt
Baixar o app pandoc_3.5.exe e instalar no windows
Link pandoc -> https://github.com/jgm/pandoc/releases/download/3.5/pandoc-3.5-windows-x86_64.msi

Baixa o app AutoHotkey 1.1 e instalar no windows
Link versao 1.1 zip -> https://www.autohotkey.com/download/1.1/AutoHotkey_1.1.37.02_setup.exe
```
 - 2
 ```txt
Instalar python 3.12
```
 - 3
```txt
Instalar as bibliotecas dependencias conforme arquivo requirements.txt
pip install -r requirements.txt
```

- 4
```txt
Para executar o arquivo no vscode

configure a chamada no final arquivo main.py
assim...

main(file_base.rtf, pgr_modelo.docx, pgr_destino.docx)
```

- 5
```txt
Para executar o arquivo no via linha de comando
por exemplo via script bat

configure a chamada no final arquivo main.py
assim...

main(sys.argv[1], sys.argv[2], sys.argv[3])

execute a linha de comando assim...
path\python.exe main.py file_base.rtf pgr_modelo.docx pgr_destino.docx
```

