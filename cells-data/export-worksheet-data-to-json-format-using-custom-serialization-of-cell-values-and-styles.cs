// Title: How to export Excel worksheet data, including cell values and styles, to JSON with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export worksheet to JSON with cell formatting | Save Excel cell styles and values as JSON using Aspose.Cells library | Serialize Excel workbook rows and cell styles to JSON in .NET | Aspose.Cells convert used range to JSON including font and background colors | Export non‑empty Excel cells to JSON file with Aspose.Cells C#
// Tags: Aspose.Cells worksheet JSON export with styles | C# serialize Excel cell formatting to JSON | Excel cell value and style JSON representation | custom POCO model for Excel to JSON conversion | Aspose.Cells used range iteration for JSON output

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExport
{
    // Represents the style information we want to serialize.
    // The sample loads an Excel workbook with Aspose.Cells, iterates each worksheet's used range, captures non‑empty cells' addresses, raw values, and style details (font name, size, bold, italic, font color, background color) into custom POCO classes, serializes the collected worksheet, row, and cell data to an indented JSON string using System.Text.Json, and writes the JSON to a specified file, creating the output directory if necessary.
    public class CellStyleInfo
    {
        public string FontName { get; set; } = string.Empty;
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string? FontColor { get; set; }          // Hex string, e.g., "#FF0000"
        public string? BackgroundColor { get; set; }    // Hex string
    }

    // Represents a single cell with its address, value and style.
    public class CellData
    {
        public string Address { get; set; } = string.Empty;    // e.g., "A1"
        public object? Value { get; set; }                     // Raw value (string, double, DateTime, etc.)
        public CellStyleInfo Style { get; set; } = null!;
    }

    // Represents a row of cells.
    public class RowData
    {
        public int RowIndex { get; set; }
        public List<CellData> Cells { get; set; } = new List<CellData>();
    }

    // Represents a worksheet with its name and rows.
    public class WorksheetData
    {
        public string Name { get; set; } = string.Empty;
        public List<RowData> Rows { get; set; } = new List<RowData>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input Excel and output JSON.
            string excelPath = @"C:\Data\Sample.xlsx";
            string jsonPath = @"C:\Data\Sample.json";

            try
            {
                // Verify that the source Excel file exists.
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                // Load the workbook using Aspose.Cells.
                Workbook workbook = new Workbook(excelPath);

                // Container for all worksheets' data.
                List<WorksheetData> workbookData = new List<WorksheetData>();

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    WorksheetData wsData = new WorksheetData { Name = sheet.Name };

                    // Determine the used range to avoid iterating over empty cells.
                    AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                    // Calculate start/end rows and columns.
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Loop through rows in the used range.
                    for (int row = startRow; row <= endRow; row++)
                    {
                        RowData rowData = new RowData { RowIndex = row };

                        // Loop through columns in the used range.
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            Style cellStyle = cell.GetStyle();

                            // Skip cells that have no value.
                            if (cell.Value == null)
                                continue;

                            // Build the custom style object.
                            CellStyleInfo styleInfo = new CellStyleInfo
                            {
                                FontName = cellStyle.Font.Name,
                                FontSize = cellStyle.Font.Size,
                                Bold = cellStyle.Font.IsBold,
                                Italic = cellStyle.Font.IsItalic,
                                FontColor = ColorToHex(cellStyle.Font.Color),
                                BackgroundColor = ColorToHex(cellStyle.ForegroundColor)
                            };

                            // Create the cell data object.
                            CellData cellData = new CellData
                            {
                                Address = cell.Name,
                                Value = cell.Value,
                                Style = styleInfo
                            };

                            rowData.Cells.Add(cellData);
                        }

                        // Add the row only if it contains at least one cell.
                        if (rowData.Cells.Count > 0)
                            wsData.Rows.Add(rowData);
                    }

                    workbookData.Add(wsData);
                }

                // Serialize the workbook data to JSON with indentation.
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(workbookData, jsonOptions);

                // Ensure the output directory exists.
                string? outputDir = Path.GetDirectoryName(jsonPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to the output file.
                File.WriteAllText(jsonPath, json);

                Console.WriteLine($"Worksheet data exported to JSON at: {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to convert System.Drawing.Color to a hex string.
        private static string? ColorToHex(Color color)
        {
            // If the color is empty (i.e., not set), return null.
            if (color.IsEmpty)
                return null;

            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}
