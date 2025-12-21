# LibreOffice Basic Macros

A small collection of LibreOffice/OpenOffice **StarBasic** macros organized as plain `.bas` files for easy version control.

## Modules

### 🧮 Module1_Calc.bas
Utility routines for **Calc** spreadsheets.
- Automates repetitive data formatting and recalculation tasks.
- Includes examples of event-driven macros (`SheetChanged`, `BeforeSave`).
- Demonstrates use of `ThisComponent` to manipulate active sheets and cell ranges.

### 🖋️ Module2_Draw.bas
Macros targeting **Draw** documents.
- Handles page and shape management.
- Example functions for resizing or repositioning shapes based on naming conventions.
- Demonstrates traversing the Draw document’s `Pages` and `Shapes` collections.

### 📄 Module3_Writer.bas
Macros designed for **Writer** documents.
- Automates document cleanup and formatting.
- Provides functions for inserting boilerplate content or adjusting paragraph styles.
- Shows use of the `Text` and `Cursor` objects for text manipulation.
> Location in LibreOffice profile (for reference): `.../user/basic/Standard/`

### 💻 strip_bas_from_xba.sh

- Extracts StarBasic from .xba XML into plain .bas (decodes XML entities).
- Usage: run in a directory containing *.xba:

```bash
./strip_bas_from_xba.sh
```

* *Output: creates ModuleName.bas next to each ModuleName.xba.*
* *Note: Works when code is inside `<source>…</source>`. If your modules embed code directly in `<script:module>…</script:module>`, this script will produce empty files.*

---

## Quick Start

### Option A — Copy into your profile
- **Linux:** copy `.bas` files to  
  `~/.config/libreoffice/4/user/basic/Standard/`
- **Windows:** copy to  
  `C:\Users\<you>\AppData\Roaming\LibreOffice\4\user\basic\Standard\`

Restart LibreOffice, then run via: **Tools → Macros → Run Macro…**

### Option B — Import via Macro Organizer
1. **Tools → Macros → Organize Macros → LibreOffice Basic…**
2. Click **Organizer… → Libraries** tab.
3. Select *Standard* → **Modules… → New → Import**
4. Choose the `.bas` files.

(yes, I did get chatgpt to write the readme...)
