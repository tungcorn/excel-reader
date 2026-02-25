# ExcelDumperTool

## About

**ExcelDumperTool** is a specialized utility designed to act as a "Translation Layer" between binary Excel test case files and AI-driven automation workflows. 

In modern automation testing, AI agents often need to read and understand existing test cases stored in spreadsheets to generate corresponding test scripts (e.g., C# Selenium classes). However, `.xlsx` files are binary formats that AI cannot read natively. This tool solves that problem by converting complex Excel data into a clean, structured, and parseable plain-text stream.

### Key Value Propositions:
- **AI-Ready Output:** Formats cell data into a readable array-like structure (`[Value]`) that LLMs can easily tokenize and analyze.
- **Structural Integrity:** Preserves critical formatting like manual line breaks (`Alt+Enter`) by encoding them as `\n`.
- **Intelligent Noise Reduction:** Automatically trims trailing empty columns and skips non-data rows to minimize token usage and focus on actual test content.
- **Dynamic Identification:** Supports accessing sheets by name or index, allowing for highly flexible automated pipelines.

---

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download) (Version 8.0 or compatible)

## How to Run

You can run this tool using the .NET CLI. The tool requires two arguments:
1. **Absolute path** to the Excel file.
2. **Sheet identifier**, which can be either the `Index` (0-based) or a `Keyword/Name` contained within the sheet's name.

### Syntax

```bash
dotnet run --project <path-to-ExcelDumperTool> -- <path_to_excel_file> <sheet_index_or_keyword>
```

---

### Examples

Assuming your terminal is working within the `ExcelDumperTool` directory:

#### 1. Run by Sheet Keyword/Name
To read a sheet that contains the name "Thêm hàng hóa", and print the result directly to your terminal:
```bash
dotnet run -- "D:\path\to\your\file.xlsx" "Thêm hàng hóa"
```

#### 2. Run by Sheet Index
To read the very first sheet in the Excel file (Index `0`), and print the result to the terminal:
```bash
dotnet run -- "D:\path\to\your\file.xlsx" 0
```

#### 3. Export Output to a Text File
If you want to save the dumped data into a text file instead of printing it to the console, use the `>` redirect operator:
```bash
dotnet run -- "D:\path\to\your\file.xlsx" 0 > output.txt
```

#### 4. Run From a Different Directory
If your terminal is located outside the `ExcelDumperTool` folder (e.g., in the parent project directory), you must specify the project path:
```bash
dotnet run --project ".\ExcelDumperTool" -- "D:\path\to\your\file.xlsx" 0
```

## Output Format
The tool will automatically detect the used range of columns and rows dynamically. It will also strip empty trailing columns. Down-lines (Alt+Enter) within Excel cells are preserved and represented as `\n` characters in the output.

Example output:
```text
--- DUMPING SHEET: Sheet1 ---
R001:	[Header 1]	[Header 2]
R002:	[Data A]	[Data B\nMulti-line here]
```
