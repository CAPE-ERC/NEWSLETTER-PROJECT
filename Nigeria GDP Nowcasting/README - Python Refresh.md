# Nigeria GDP nowcast refresh

The complete setup, input-data dictionary, run instructions, output paths, validation
steps, and troubleshooting guide are maintained in the project-level README:

```text
..\README.md
```

From `NEWSLETTER-PROJECT`, run:

```powershell
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

Recommended bundled runtime:

```powershell
& "C:\Users\eakan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

The GDP workbook remains in this folder. Monthly newsletters are written to:

```text
..\CAPE ERC Economic Newsletter\YYYY\Month Newsletter
```

