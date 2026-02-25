# ExcelDumperTool

ExcelDumperTool is a lightweight C# console application designed to read Excel (`.xlsx`) files and dump the contents of a specific worksheet to the console (or a text file) in a structured, plain-text format. 

This tool is particularly useful for AI agents or automated pipelines that need to read and parse Excel files without requiring complex programmatic interactions with Excel interop or UI applications.

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download) (Version 8.0 or compatible)

## How to Run

You can run this tool using the .NET CLI. The tool requires two arguments:
1. **Absolute path** to the Excel file.
2. **Sheet identifier**, which can be either the `Index` (0-based) or a `Keyword/Name` contained within the sheet's name.

### Syntax

```bash
dotnet run --project "<ABSOLUTE_PATH_TO_EXCEL_DUMPER_TOOL>" -- <path_to_excel_file> <sheet_index_or_keyword>
```

---

### Examples

Assuming your terminal is working within the `ExcelDumperTool` directory:

#### 1. Run by Sheet Keyword/Name
To read a sheet that contains the name "Thêm nhóm hàng hóa" (Add Product Group), pass a keyword:
```bash
dotnet run -- product-group-management.xlsx "Thêm nhóm"
```

#### 2. Run by Sheet Index
To read a specific sheet by its 0-based index instead of its name:
```bash
# Read the first sheet
dotnet run -- product-group-management.xlsx 0

# Read the second sheet
dotnet run -- product-group-management.xlsx 1
```

#### 3. Run All Sheets (NEW)
To automatically iterate through and dump **all sheets** in the Excel file, use the keyword **`all`**:
```bash
dotnet run -- product-group-management.xlsx all
```
*Note: The output of each sheet will be separated by the header `--- DUMPING SHEET: Sheet Name ---` for easier AI parsing.*


#### 4. Export Output to a Text File
If you want to save the dumped data into a text file instead of printing it to the console, use the `>` redirect operator:
```bash
dotnet run -- "D:\path\to\your\file.xlsx" 0 > output.txt
```

#### 5. Run From a Different Directory
If your terminal is located outside the `ExcelDumperTool` folder (e.g., in the parent project directory), you must specify the project path:
```bash
dotnet run --project ".\ExcelDumperTool" -- "D:\path\to\your\file.xlsx" 0
```

#### 6. Comprehensive Example (Recommended for AI Workflows)
If you are working inside your main test project directory (e.g. `TestProductGroup`) and want to dump *all* sheets into a single text file using an absolute path, use the full project path:
```bash
dotnet run --project "<ABSOLUTE_PATH_TO_EXCEL_DUMPER_TOOL>" -- "<ABSOLUTE_PATH_TO_EXCEL_FILE>" all > all_sheets_output.txt
```
## Output Format
The tool will automatically detect the used range of columns and rows dynamically. It will also strip empty trailing columns. Down-lines (Alt+Enter) within Excel cells are preserved and represented as `\n` characters in the output.

Example output:
```text
--- DUMPING SHEET: Sheet1 ---
R001:	[Header 1]	[Header 2]
R002:	[Data A]	[Data B\nMulti-line here]
```
