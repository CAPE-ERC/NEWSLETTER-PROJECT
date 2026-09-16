# Nigeria GDP nowcast: Python refresh

Save your monthly inputs in `Main Data/MultipleFrequencyModelData.xlsx`. Close the output nowcast workbook and newsletter before running. Do not edit the master while the refresh runs; unsaved Excel edits are not read.

From the `NEWSLETTER-PROJECT` folder, run:

```text
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

If Python is not on PATH, use the installed runtime (this terminal command launches Python, not a PowerShell script):

```powershell
& "C:\Users\eakan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" ".\GDP_Nowcast_and_Newsletter.py" 2026-10
```

Omitting the month uses the current calendar month. First check paths without generating anything:

```text
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --dry-run
```

The launcher estimates the existing model, builds and validates a freshly saved workbook, then publishes the newsletter. It does not change the equation, write the master workbook, or create master-workbook backups. The existing Node.js workbook renderer remains an internal dependency; no `.ps1` file is called. Use a Python installation with the dependencies required by the existing analysis/publishing scripts; the bundled runtime above already serves the existing workflow.

The month selects the newsletter edition, not a forced GDP target quarter or a historical data cutoff. The model still determines its quarter from available source data. One, two or three available monthly observations follow the existing model rules. The existing publication range and completeness rules are unchanged.

The clean output name is `CAPE Economic Performance and Prospect Bulletin October 2026.docx`, in `NEWSLETTER-PROJECT/CAPE ERC Economic Newsletter/2026/October Newsletter`. Each edition is stored under its year, and existing output for that month is refreshed in place. Template priority is an explicit template, the existing report, a monthly `- Template.docx`, then the previous month's clean report (including across a year boundary).

Optional arguments:

```text
python ".\GDP_Nowcast_and_Newsletter.py" 2026-10 --newsletter-template "C:\path\October Template.docx" --plots-folder "C:\path\Plots"
```

Without a plots argument, the existing publisher uses the edition's `Plots` folder. Follow its existing figure naming/mapping conventions; this launcher does not change them. Review all charts, dates and editorial text before distribution: refreshing an older template is not a guarantee that every section has been newly researched.

On failure, the launcher stops and reports the error. Model/build validation failure prevents newsletter publication. A renderer teardown error is tolerated only when a fresh workbook exists and the existing verifier passes. The original `.ps1` is retained for compatibility but is not needed for this Python entry point. Avoid running both launchers simultaneously.
