using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ExcelDumper
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExcelDumper <path_to_excel_file> <sheet_index_or_keyword>");
                return;
            }

            string filePath = args[0];
            string sheetIdentifier = args[1];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found at '{filePath}'");
                return;
            }

            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));

                var sheetsToDump = new System.Collections.Generic.List<ExcelWorksheet>();

                if (string.Equals(sheetIdentifier, "all", StringComparison.OrdinalIgnoreCase))
                {
                    sheetsToDump.AddRange(package.Workbook.Worksheets);
                }
                else
                {
                    ExcelWorksheet targetSheet = null;
                    if (int.TryParse(sheetIdentifier, out int sheetIndex) && sheetIndex >= 0 && sheetIndex < package.Workbook.Worksheets.Count)
                    {
                        targetSheet = package.Workbook.Worksheets[sheetIndex];
                    }
                    else
                    {
                        targetSheet = package.Workbook.Worksheets.FirstOrDefault(
                            ws => !string.IsNullOrWhiteSpace(ws.Name) &&
                                  ws.Name.Contains(sheetIdentifier, StringComparison.OrdinalIgnoreCase));
                    }

                    if (targetSheet == null)
                    {
                        Console.WriteLine($"Error: Sheet '{sheetIdentifier}' not found.");
                        Console.WriteLine("Available sheets: " + string.Join(", ", package.Workbook.Worksheets.Select((ws, i) => $"[{i}] '{ws.Name}'")));
                        return;
                    }
                    
                    sheetsToDump.Add(targetSheet);
                }

                bool isFirst = true;
                foreach (var sheet in sheetsToDump)
                {
                    if (!isFirst) Console.WriteLine();
                    isFirst = false;

                    Console.WriteLine($"--- DUMPING SHEET: {sheet.Name} ---");
                    
                    int maxRow = sheet.Dimension?.End.Row ?? 1;
                    int maxCol = sheet.Dimension?.End.Column ?? 1;

                    // Tự động sử dụng đúng số cột thực tế có kích thước trong Excel file
                    // Khống chế an toàn tối đa 50 cột để tránh việc file Excel bị đục chuẩn format (chứa khoảng trắng vô tận)
                    maxCol = Math.Min(maxCol, 50); 

                    for (int row = 1; row <= maxRow; row++)
                    {
                        System.Collections.Generic.List<string> rowCells = new System.Collections.Generic.List<string>();
                        bool hasDataInRow = false;

                        for (int col = 1; col <= maxCol; col++)
                        {
                            var cellValue = sheet.Cells[row, col].Text?.Trim() ?? "";
                            
                            // Xử lý cả rớt dòng thực sự (\n) và rớt dòng bị escape (\\n hoặc \r) của EPPlus
                            cellValue = Regex.Replace(cellValue, @"\r\n?|\n|\\n|\\r", "\\n");
                            
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                hasDataInRow = true;
                            }
                            
                            rowCells.Add($"[{cellValue}]");
                        }

                        if (hasDataInRow)
                        {
                            // Lọc bỏ các cột trống vô nghĩa ở tận cùng bên phải của mỗi dòng
                            int lastNonEmptyIndex = rowCells.FindLastIndex(c => c != "[]");
                            if (lastNonEmptyIndex >= 0)
                            {
                                var trimmedCells = rowCells.Take(lastNonEmptyIndex + 1);
                                Console.WriteLine($"R{row:D3}:\t{string.Join("\t", trimmedCells)}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
            }
        }
    }
}
