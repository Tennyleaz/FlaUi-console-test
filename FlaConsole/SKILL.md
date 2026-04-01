---
name: computer-use
description: Control windows applications using command line tools.
metadata:
  {
    "openclaw":
      {
        "emoji": "💻",
        "requires": { },
      },
  }
---
 
# skill: Control Windows Apps via FlaConsole.exe (UI Automation)

This skill controls Windows desktop applications by calling a local command-line tool: `FlaConsole.exe`.

**Important:** `FlaConsole.exe` is located in the same directory as this `SKILL.md`. Always invoke it via **relative path**:

- Windows PowerShell / CMD: `.\FlaConsole.exe ...`

---

## Tool Overview

Supported commands (summary):

- `flaui list [--json]`  
  List top-level windows (PID, handle, process, title)

- `flaui tree --pid <pid> [--depth <n>] [--max-items <n>] [--json]`  
  Print UI Automation tree for a process/window

- `flaui find --pid <pid> --path <i1,i2,...> [--require-match] [--fallback] [--json]`  
  Resolve a specific element by a *path* and show its properties

- `flaui click --pid <pid> --path <i1,i2,...> [--dry-run] [--json]`  
  Click a UI element identified by *path*

- `flaui subtree --pid <pid> --path <i1,i2,...> [--depth <n>] [--max-items <n>] [--json]`  
  Show only a subtree starting at the specified element

### Path Format

A UI element is addressed by a comma-separated list of indexes:

- Example: `0,1,1,1,7,3`
- Index `0` is always the root.
- Each index selects the *N-th child* at that level as shown in `tree` output.

---

## Standard Operating Procedure (SOP)

### 1) Identify the target app/window PID

Use:

```
.\FlaConsole.exe list --json
```

Pick the PID that matches the target window title/process.

**Tip:** If multiple windows match, prefer the one whose title best matches the task.

---

### 2) Inspect the UI tree to find relevant controls

Use:

```powershell
.\FlaConsole.exe tree --pid <pid>
```

If output is too large, limit it:

```powershell
.\FlaConsole.exe tree --pid <pid> --depth 6 --max-items 200
```

Locate the desired control (e.g., a button or menu item) and copy its `path=...`.

---

### 3) (Optional but recommended) Verify the element before clicking

Use:

```powershell
.\FlaConsole.exe find --pid <pid> --path <path>
```

Confirm:
- `controlType` matches intent (Button, MenuItem, etc.)
- `name` / `automationId` looks correct
- `enabled: True`

---

### 4) Click the control

Use:

```powershell
.\FlaConsole.exe click --pid <pid> --path <path>
```

For safety/testing:

```powershell
.\FlaConsole.exe click --pid <pid> --path <path> --dry-run
```

---

## Troubleshooting & Reliability Rules

1. **If PID changes** (app restarted, window reopened), re-run `list` to get the new PID.
2. **If paths change** (UI updated, different app mode), re-run `tree` and re-locate the control.
3. Prefer controls with stable identifiers in the tree output:
   - `automationId` (best)
   - `name`
4. If the full `tree` is too big:
   - Find an approximate parent first, then use `subtree` from that parent path:
     ```powershell
     .\FlaConsole.exe subtree --pid <pid> --path 0,1,1 --depth 6 --max-items 200
     ```
5. If clicking does nothing:
   - Verify `enabled: True` via `find`
   - Re-check you selected the correct window PID/title

6. Use `--json` so output could be parsed better.
---

## Example: Operating Windows Calculator (calc)

### Step 1: Find PID

```powershell
.\FlaConsole.exe list
```

Example output:

```
[pid]   [handle]        [process]       [title]
16120   917554  ApplicationFrameHost    小算盤
```

So `小算盤` (Calculator in zh-tw locale) PID is `16120`.

### Step 2: Inspect tree

```powershell
.\FlaConsole.exe tree --pid 16120
```

Find the number pad button "三" (three in zh-tw locale) with path:

- `"三"` (id=`num3Button`) at `path=0,1,1,1,7,3`

### Step 3: Click "3"

```powershell
.\FlaConsole.exe click --pid 16120 --path 0,1,1,1,7,3
```

Expected result includes confirmation of the found element and “Click executed.”

---

## Minimal Command Templates (Copy/Paste)

```powershell
# list windows
.\FlaConsole.exe list --json

# view UI tree (optionally limit size)
.\FlaConsole.exe tree --pid <pid> --depth 8 --max-items 200

# verify a specific element
.\FlaConsole.exe find --pid <pid> --path 0,1,2,3

# click an element
.\FlaConsole.exe click --pid <pid> --path 0,1,2,3
```

