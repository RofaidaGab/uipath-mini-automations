# UiPath Mini Automations

A collection of small, self-contained UiPath workflows — each one demonstrates a different automation pattern (web login, file organizing, web scraping, PDF/Excel data extraction, form filling, mail triage, and Word/Excel reporting). Built as practice projects in UiPath Studio (Studio version `26.0.199.0`, C#/modern project format).

Every workflow can be run independently from Studio — there is no single orchestrated "run everything" pipeline (`Main.xaml` is an empty placeholder entry point required by the `.uiproj` format).

## Repository structure

```
uipath-mini-automations/
├── Credential.xaml          # Orchestrator Asset login demo (Coursera)
├── Excel Automation.xaml    # Read → sort → rewrite an Excel report
├── File_Org.xaml            # Sort files into folders by extension
├── Form_Filler.xaml         # Excel-driven web form filler (RPA Challenge)
├── Mail_Automation.xaml     # Gmail inbox triage (Integration Service)
├── PDF_Auto.xaml            # Bulk PDF invoice data extraction → Excel
├── Report generator.xaml    # Excel data → Word template → PDF report
├── Web_Scraping.xaml        # Product table scraper → Excel
├── Main.xaml                # Empty placeholder entry point
├── project.json             # UiPath project manifest
├── entry-points.json        # UiPath entry point manifest
├── project.uiproj           # Studio project pointer file
├── .gitignore
├── Data/
│   ├── Input/                # <- put your own input files here (see table)
│   ├── Templates/             # <- put Sales_Report_Template.docx here
│   └── Output/                # <- generated files land here (git-ignored)
└── README.md
```

> **Note:** The `Data/` folder is not part of the original UiPath project — it was added when preparing this repo for GitHub so no personal file paths (e.g. `C:\Users\<name>\...`) would be published. All workflow file paths below have been updated to point inside `Data/` (relative to the project root, which is how UiPath resolves them at runtime). Generated output files are git-ignored — see [Setup](#setup) for what to add locally.

## Workflow summary

| Workflow | Purpose | Key activities used | Input | Output |
|---|---|---|---|---|
| `Credential.xaml` | Logs into Coursera using a securely stored Orchestrator credential asset instead of a hardcoded password. | `Get Robot Credential`, `Application Card` (Edge), `Type Into`, `Click` | Orchestrator Asset `Website_Credentials` (Shared folder) | Logged-in Coursera session |
| `File_Org.xaml` | Watches a user-selected folder and sorts files into sub-folders by extension (`.pdf`, `.xlsx`, etc.), with try/catch error handling. | `Browse For Folder`, `For Each File in Folder`, `Switch`, `Move File`, `Try Catch` | Any folder (chosen at runtime via dialog) | Files moved into `pdfFolder` / `excelFolder` sub-paths |
| `Web_Scraping.xaml` | Scrapes a Trendyol electronics category listing table and writes the results to Excel. | `Application Card` (Edge), `Extract Table Data`, `Excel Process Scope`, `Write Range` | Live Trendyol page (no local input) | `Data/Output/Products.xlsx` |
| `PDF_Auto.xaml` | Loops through a folder of PDF invoices, reads each with OCR/text extraction, and logs Invoice Number / Date / Total to Excel. | `For Each File in Folder`, `Read PDF Text`, `Build Data Table`, `Add Data Row`, `Excel Process Scope` | `Data/Input/Invoices/*.pdf` | `Data/Output/InvoiceLog.xlsx` |
| `Form_Filler.xaml` | Reads rows from an Excel sheet and types each field (Name, Company, Email, Phone, Address, Role) into the public RPA Challenge web form. | `Excel Process Scope`, `Read Range`, `For Each Row`, `Switch`, `Type Into` | `Data/Input/form.xlsx` | Submitted web form on [rpachallenge.com](https://rpachallenge.com/) |
| `Mail_Automation.xaml` | Polls a Gmail inbox (via UiPath Integration Service) and routes unread mail based on sender: known senders get moved to Trash / auto-replied / marked unread; everything else is logged as uncategorized. | `Get Email List`, `For Each`, `Switch`, `Move Email`, `Reply to Email`, `Mark Email as Read/Unread` | Gmail inbox (via connected Integration Service account) | Mailbox actions (move/reply/mark) |
| `Excel Automation.xaml` | Reads a range from a source workbook, sorts it (via an Invoke Code step), and writes the sorted table back out. | `Excel Process Scope`, `Read Range`, `Sort Data Table`, `Invoke Code`, `Write Range` | `Data/Input/AppOneReport.xlsx` | Same workbook, sorted in place |
| `Report generator.xaml` | Reads sales data from Excel, fills a Word template (bookmarks + placeholder text like `{{TotalTransactions}}`), inserts the data table, and exports the result as PDF. | `Excel Process Scope`, `Read Range`, `Word Application Scope`, `Replace Text`, `Insert DataTable in Document`, `Save Document as PDF` | `Data/Input/Input_Sales_Data.xlsx`, `Data/Templates/Sales_Report_Template.docx` | `Data/Output/Sales_Report_Template.pdf` |

## Setup

1. Clone the repo and open `project.json` in **UiPath Studio** (2024.10+ / the C#-based "Windows" target — this project was authored on Studio `26.0.199.0`).
2. Let Studio restore the NuGet dependencies listed in `project.json` (Excel, GSuite/Gmail, Integration Service, Microsoft 365, PDF, UI Automation, Word activities).
3. Create the local data folders and drop in your own files (these are git-ignored so your personal data never gets committed):

   | File to add | Goes in | Used by |
   |---|---|---|
   | Any `.pdf` invoices | `Data/Input/Invoices/` | `PDF_Auto.xaml` |
   | `form.xlsx` (name/company/email/phone/address/role columns) | `Data/Input/` | `Form_Filler.xaml` |
   | `AppOneReport.xlsx` | `Data/Input/` | `Excel Automation.xaml` |
   | `Input_Sales_Data.xlsx` | `Data/Input/` | `Report generator.xaml` |
   | `Sales_Report_Template.docx` (with `{{TotalTransactions}}` / `{{TopRep}}` placeholders and a `SalesTableBookmark` bookmark) | `Data/Templates/` | `Report generator.xaml` |

4. **Orchestrator asset:** `Credential.xaml` expects a Credential-type asset named `Website_Credentials` in the `Shared` folder of your Orchestrator tenant. Create it (or point `AssetName`/`FolderPath` at your own) before running.
5. **Gmail / Integration Service:** `Mail_Automation.xaml` uses a UiPath Integration Service **Gmail** connection. Add your own Gmail connection in Integration Service, then re-point the activities at your `ConnectionId` (Studio will prompt you to reconnect on first open). The workflow's own-account routing case is set to the placeholder `your-email@gmail.com` — change it to your address in the `Switch` on `strSender`.

## Running a workflow

1. Open the repo folder as a UiPath project in Studio.
2. In the **Project** panel, double-click the `.xaml` file you want to run (e.g. `Web_Scraping.xaml`).
3. Press **F5** / **Run** — Studio runs whichever file is open as its own entry point (you don't need to run `Main.xaml`, which is intentionally empty).
4. Some workflows launch a browser via an **Application Card**; on first run in a new environment, Studio's Object Repository/selector may need to be re-validated (UI Automation selectors are tied to the recorded app version/DOM).

## Security notes

- No plaintext passwords are stored anywhere in this repo — `Credential.xaml` pulls credentials from an Orchestrator asset at runtime, and `Mail_Automation.xaml` authenticates via UiPath Integration Service (OAuth), not stored passwords.
- All local `C:\Users\...` file paths from the original project have been replaced with relative `Data/...` paths.
- The personal Gmail address originally hardcoded into `Mail_Automation.xaml`'s routing logic has been replaced with a placeholder (`your-email@gmail.com`) — update it to your own address before use.
