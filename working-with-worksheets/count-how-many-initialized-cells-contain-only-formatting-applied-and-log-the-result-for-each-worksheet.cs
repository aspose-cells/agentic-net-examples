// Title: Count formatting‑only cells (no values) in each worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# count cells that have style but no data in a worksheet | How to get number of empty cells with custom formatting using Aspose.Cells .NET | C# iterate worksheets and find cells with only formatting applied | Determine count of styled empty cells per sheet with Aspose.Cells API | Aspose.Cells example to log formatting‑only cells in each worksheet | Calculate max styled row and column in Aspose.Cells to scan formatting only cells
// Tags: count formatting‑only cells Aspose.Cells | detect styled empty cells .NET | worksheet iteration style detection Aspose.Cells | max styled row calculation Aspose.Cells | create demo workbook formatting only Aspose.Cells | cell style comparison default Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace FormattingOnlyCellCounter
{
    // The C# program loads an Excel workbook, determines scanning bounds by combining the maximum data row/column with the highest row/column that contains a non‑default style, and iterates each worksheet to count cells whose value is null or empty while their style differs from the workbook's default style. It logs the count per sheet, uses helper methods to compute the maximum styled row and column, and includes a demo workbook creator that adds sample cells with only formatting, only values, and both for testing.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // Input workbook handling
            // ------------------------------------------------------------
            string inputPath;
            if (args.Length > 0 && File.Exists(args[0]))
            {
                inputPath = args[0];
            }
            else
            {
                // Create a demo workbook if no valid file is supplied
                inputPath = "DemoWorkbook.xlsx";
                CreateDemoWorkbook(inputPath);
                Console.WriteLine($"No valid input file supplied. A demo workbook has been created at '{Path.GetFullPath(inputPath)}'.");
            }

            try
            {
                // Load the workbook (ensure the file exists)
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Workbook file not found: {inputPath}");

                Workbook workbook = new Workbook(inputPath);
                Style defaultStyle = workbook.DefaultStyle;

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine scanning bounds (data + styled cells)
                    int maxRow = Math.Max(cells.MaxDataRow, GetMaxStyledRow(cells, defaultStyle));
                    int maxCol = Math.Max(cells.MaxDataColumn, GetMaxStyledColumn(cells, defaultStyle));

                    // Adjust for empty worksheets
                    if (maxRow < 0) maxRow = 0;
                    if (maxCol < 0) maxCol = 0;

                    int formattingOnlyCount = 0;

                    // Scan the determined range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Determine if the cell has no value
                            bool hasNoValue = cell.Value == null || string.IsNullOrEmpty(cell.StringValue);

                            // Determine if the cell style differs from the default style
                            bool hasCustomStyle = !cell.GetStyle().Equals(defaultStyle);

                            // Count cells that have formatting only
                            if (hasNoValue && hasCustomStyle)
                            {
                                formattingOnlyCount++;
                            }
                        }
                    }

                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {formattingOnlyCount} cell(s) with only formatting applied.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        // Returns the highest row index that contains a non‑default style
        private static int GetMaxStyledRow(Cells cells, Style defaultStyle)
        {
            int maxRow = -1;
            int dataMaxRow = cells.MaxDataRow;
            int dataMaxCol = cells.MaxDataColumn;

            if (dataMaxRow < 0) dataMaxRow = 0;
            if (dataMaxCol < 0) dataMaxCol = 0;

            for (int r = 0; r <= dataMaxRow; r++)
            {
                for (int c = 0; c <= dataMaxCol; c++)
                {
                    Cell cell = cells[r, c];
                    if (!cell.GetStyle().Equals(defaultStyle))
                    {
                        if (r > maxRow) maxRow = r;
                    }
                }
            }

            return maxRow;
        }

        // Returns the highest column index that contains a non‑default style
        private static int GetMaxStyledColumn(Cells cells, Style defaultStyle)
        {
            int maxCol = -1;
            int dataMaxRow = cells.MaxDataRow;
            int dataMaxCol = cells.MaxDataColumn;

            if (dataMaxRow < 0) dataMaxRow = 0;
            if (dataMaxCol < 0) dataMaxCol = 0;

            for (int r = 0; r <= dataMaxRow; r++)
            {
                for (int c = 0; c <= dataMaxCol; c++)
                {
                    Cell cell = cells[r, c];
                    if (!cell.GetStyle().Equals(defaultStyle))
                    {
                        if (c > maxCol) maxCol = c;
                    }
                }
            }

            return maxCol;
        }

        // Creates a demo workbook with cells that have only formatting applied
        private static void CreateDemoWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Cell A1 – value only
            cells["A1"].PutValue("Data");

            // Cell B2 – formatting only (background color)
            cells["B2"].SetStyle(CreateCustomStyle(wb, Color.Yellow));

            // Cell C3 – formatting only (font bold)
            cells["C3"].SetStyle(CreateCustomStyle(wb, null, true));

            // Cell D4 – both value and formatting (should NOT be counted)
            cells["D4"].PutValue(123);
            cells["D4"].SetStyle(CreateCustomStyle(wb, Color.LightGreen, true));

            wb.Save(path, SaveFormat.Xlsx);
        }

        // Helper to create a style with optional background color and bold font
        private static Style CreateCustomStyle(Workbook wb, Color? backColor = null, bool bold = false)
        {
            Style style = wb.CreateStyle();
            if (backColor.HasValue)
            {
                style.ForegroundColor = backColor.Value;
                style.Pattern = BackgroundType.Solid;
            }
            if (bold)
            {
                style.Font.IsBold = true;
            }
            return style;
        }
    }
}
