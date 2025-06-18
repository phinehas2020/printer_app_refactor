# Printer App Refactor

This project is a refactored version of the original monolithic `Currentv2.py` printer app. The codebase is now organized into logical modules for maintainability and clarity.

## Structure

- `app/` - App entry point and Flask app creation
- `config/` - Configuration management (load/save config, defaults)
- `shopify_api/` - Shopify API interaction (GraphQL, inventory)
- `label_utils/` - Label creation, image manipulation, barcode handling
- `printer_utils/` - Printer communication (Brother, Epson, etc.)
- `spreadsheet_utils/` - Spreadsheet loading, parsing, and utilities
- `routes/` - Flask route handlers, grouped by functionality:
  - `index.py` - Main page and admin config
  - `products.py` - Product endpoints
  - `printing.py` - Print label, print status, cancel print
  - `inventory.py` - Inventory status and update
  - `admin.py` - Admin login, config, product management
  - `debug.py` - Debug log endpoint

## How to Run

1. Install dependencies (see requirements.txt)
2. Set up your configuration (see `config/config.py`)
3. Run the app:
   ```bash
   python app/app.py
   ```

## Notes
- The original monolithic file is preserved as `Currentv2.py` for reference.
- All secrets (API keys, tokens) must be set via environment variables or config files, not in code.
