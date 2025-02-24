# Conversor de Arquivos Paradox para Excel

Este projeto permite a conversão de arquivos de banco de dados Paradox (`.db`) para arquivos Excel (`.xlsx`). Ele corrige caracteres corrompidos e garante que os dados sejam exportados corretamente.

## 📌 Funcionalidades

- Seleção de arquivos Paradox (`.db`) via interface gráfica.
- Conversão automática para formato Excel (`.xlsx`).
- Correção de caracteres corrompidos.
- Tratamento de valores nulos e datas inválidas.
- Notificação sobre a conclusão da conversão.

## 📋 Requisitos

Antes de executar o script, certifique-se de ter as seguintes bibliotecas instaladas:

```bash
pip install pandas tkinter pypxlib openpyxl
```

## 🚀 Como Usar

1. Execute o script `python`:

   ```bash
   python script.py
   ```

2. Uma janela será aberta para selecionar o arquivo `.db`.
3. O script processará os dados e criará um arquivo `.xlsx` na mesma pasta do arquivo original.
4. Uma mensagem será exibida ao final informando o sucesso da conversão.

## 🛠 Configuração e Personalização

Se necessário, você pode modificar o dicionário `CHAR_FIXES` no código para adicionar mais correções de caracteres.

```python
CHAR_FIXES = {
    "Ã├O": "ÇÃO",
    "Ã£": "ã",
    "Ã¡": "á",
    "Ã©": "é",
    "Ãª": "ê",
    "Ã³": "ó",
    "Ã§": "ç",
    "Ãº": "ú",
    "Ã­": "í",
    "Ãµ": "õ",
    "Ã¤": "ä",
    "Ã¶": "ö",
}
```

## 📄 Licença

Este projeto é de código aberto e pode ser modificado conforme necessário. Só solicitado referenciar o author. 

## 👦 Author

* [Patrick Lohn](https://github.com/patricklohn)
---

📌 Desenvolvido para facilitar a conversão de arquivos Paradox para Excel! 🚀

