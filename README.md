# Study Guide Watcher

Convert Markdown study notes into Word documents. Put a `.md` file in `inbox/`; the watcher creates a formatted `.docx` in `outbox/` and, on success, moves the source file to `done/`.

## Features

- A title page using the first `#` heading, or the filename if there is no title heading
- A Table of Contents field that you can update in Word
- Word heading styles for Markdown headings
- Bullet and numbered lists
- Basic pipe-delimited tables
- Shaded code blocks and italicized block quotes
- Bold and italic inline text
- Conversion results recorded in `conversion_log.txt`

## Project Layout

```text
Study Pipeline/
├── .venv/                     # Python virtual environment
├── .vscode/
│   └── tasks.json             # VS Code launch task
├── watcher.py                 # Watcher and Markdown-to-Word converter
├── start_watcher.bat          # Windows one-click launcher
├── README.md
├── inbox/                     # Add .md files here
├── outbox/                    # Converted .docx files appear here
├── done/                      # Successfully converted .md files move here
└── conversion_log.txt         # Created when the watcher logs an event
```

The `inbox/`, `outbox/`, and `done/` folders are created automatically if they do not exist.

## Run the Watcher

### VS Code

1. Open this project folder in VS Code.
2. Press **Ctrl+Shift+B** to run the default build task.
3. Leave the task terminal open while the watcher runs. Press **Ctrl+C** there to stop it.

You can also use **Ctrl+Shift+P**, then select **Tasks: Run Build Task**.

### Windows

Double-click `start_watcher.bat` in File Explorer. The batch file runs the project virtual environment's Python. Press **Ctrl+C** in its console to stop the watcher; press a key to close the console afterward.

## How It Works

- At startup, the watcher converts `.md` files already in the top level of `inbox/`.
- While running, it watches that folder for newly created `.md` files. It does not watch nested folders or later edits to files already there.
- The title is taken from the first `#` heading. If there is none, the filename is used.
- On successful conversion, the `.docx` is saved in `outbox/` with an unused name. For repeated names, outputs are numbered like `hello.docx`, `hello (1).docx`.
- The Markdown source keeps its original filename. If that filename is already in `done/`, the new copy is placed in a numbered subfolder, such as `done/hello (1)/hello.md`.
- If conversion fails, the error is logged and the source stays in `inbox/`.

The watcher waits briefly after detecting a new file so a copy can finish. Conversion time depends on the file and computer; it is not guaranteed to finish within one second.

## Markdown Supported

The converter handles a basic Markdown subset. It does not render links, images, nested lists, or other Markdown extensions.

### Headings

```markdown
# Document title
## Main section
### Subsection
#### Smaller subsection
```

The first `#` heading is used on the title page. In the document body, `##`, `###`, and `####` map to Word Heading 1, Heading 2, and Heading 3 styles. A `#` heading also appears in the body using Word's Title style.

### Lists and Inline Formatting

```markdown
- Bullet item
- Another bullet

1. Numbered item
2. Another item

This is **bold** and this is *italic*.
```

### Code, Quotes, and Rules

Use triple backticks for a fenced code block, `> ` at the start of a line for a quote, and a line containing `---` for a horizontal rule.

### Tables

```markdown
| Concept | Meaning |
|---------|---------|
| Rank 0  | Scalar  |
| Rank 1  | Vector  |
```

## First-Time Setup

Run these commands from the project folder in PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install python-docx watchdog
```

To confirm the dependencies are installed in this virtual environment:

```powershell
.\.venv\Scripts\python.exe -m pip show python-docx watchdog
```

## Word Table of Contents

The document contains a Table of Contents field. To populate or refresh it in Word, right-click the table and choose **Update Field**, then **Update entire table**.

## Troubleshooting

| Symptom | What to check |
|---------|---------------|
| `ModuleNotFoundError` for `docx` or `watchdog` | Install dependencies with `\.venv\Scripts\python.exe -m pip install python-docx watchdog` from the project folder. |
| VS Code build shortcut does not start the watcher | Confirm `.vscode/tasks.json` exists and the project folder, including `.venv/`, is open in VS Code. |
| No `.docx` appears | Check `conversion_log.txt`. If conversion failed, the source remains in `inbox/`. |
| Nothing happens after adding a file | Confirm the watcher is running and the `.md` file is directly inside `inbox/`, not a subfolder. |
| The Table of Contents is blank or outdated | In Word, right-click it and choose **Update Field** → **Update entire table**. |

## Customize the Output

Open `watcher.py` to change the title-page accent color, code-block background, quote color, or table style. For example, the table style is set by:

```python
tbl.style = 'Light Grid Accent 1'
```

Word's built-in heading styles control the body heading appearance.

## File Management

- Successfully converted Markdown files are moved from `inbox/` to `done/` without changing their filenames. Repeated source names are stored in numbered subfolders under `done/`.
- Files that fail conversion remain in `inbox/`; check the log before retrying.
- Converted Word documents are saved in `outbox/`. Repeated base filenames get numbered output names rather than replacing existing documents.
- `conversion_log.txt` can be cleared manually when the watcher is stopped.