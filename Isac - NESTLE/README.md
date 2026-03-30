# Simexpress Bot

Automação de login no https://simexpress.com.br/, navegação em Pedidos > Em Lote, inserção de pedidos / linha e download CSV.

## Passos

1. Abra PowerShell nesta pasta
2. Crie venv:
   ```powershell
   python -m venv .venv
   ```
3. Ative:
   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```
& "c:\Users\BRCavalhIs\NESTLE\Transportes CCS - Automações\Automação Simexpress\Isac\Isac\.venv\Scripts\Activate.ps1"

4. Instale:
   ```powershell
   pip install selenium webdriver-manager python-dotenv openpyxl
   ```
5. Crie `.env` com suas credenciais (exemplo abaixo)
6. Execute:
   ```powershell
   python simexpress_bot.py
   ```
   cd "c:\Users\BRCavalhIs\NESTLE\Transportes CCS - Automações\Automação Simexpress\Isac\Isac"; & ".\.venv\Scripts\Activate.ps1"; python simexpress_bot.py

## .env

```
SIMEXPRESS_USUARIO=seu_usuario (acompanhamentonespressope.log)
SIMEXPRESS_SENHA=sua_senha (acompanhamentonespressope)
DOWNLOAD_PATH=C:\Temp\simexpress
```