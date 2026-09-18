// Title: Generate a macro‑free Excel report of cells styled with the Hyperlink theme color using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# console application with Aspose.Cells that opens a workbook, scans every worksheet, and records the worksheet name and address of each cell whose font ThemeColor is Hyperlink into a new report sheet. | Enhance the Hyperlink theme color report to also capture the displayed text of each matched cell alongside its address. | Modify the solution to skip hidden worksheets and to sort the report rows alphabetically by worksheet name before saving.
// Common Searches: how to list all cells with hyperlink theme color in an Excel file using Aspose.Cells C# | Aspose.Cells generate report of cells styled with Hyperlink font color without macros | C# scan workbook for ThemeColor Hyperlink and export results to a new worksheet
// Tags: Aspose.Cells detect Hyperlink theme color cells | C# create macro‑free Excel report | iterate worksheets used range Aspose.Cells | write cell address and value to new sheet C# | auto-fit columns Aspose.Cells worksheet

using System;
using System.IO;
using Aspose.Cells;

namespace HyperlinkThemeColorReport
{
    // The example loads a macro‑free workbook, adds a dedicated report worksheet, iterates through each non‑report worksheet's used range, checks every cell's font ThemeColor for the Hyperlink value, and logs the worksheet name and cell address for matches. Columns are auto‑fitted and the workbook is saved, producing a concise report of all Hyperlink‑styled cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "HyperlinkThemeColorReport.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the source workbook (macro‑free)
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Add a new worksheet for the report
                int reportSheetIndex = sourceWorkbook.Worksheets.Add();
                Worksheet reportSheet = sourceWorkbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "HyperlinkThemeColorReport";

                // Header for the report
                reportSheet.Cells[0, 0].PutValue("Worksheet");
                reportSheet.Cells[0, 1].PutValue("Cell Address");

                int reportRow = 1; // Start writing after header

                // Iterate through all worksheets in the source workbook
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    // Skip the report sheet itself if it already exists in the collection
                    if (ws.Name == reportSheet.Name)
                        continue;

                    // Determine the used range of the worksheet
                    Aspose.Cells.Range usedRange = ws.Cells.MaxDisplayRange;
                    if (usedRange == null)
                        continue; // No used cells

                    int firstRow = usedRange.FirstRow;
                    int lastRow = usedRange.RowCount + firstRow - 1;
                    int firstColumn = usedRange.FirstColumn;
                    int lastColumn = usedRange.ColumnCount + firstColumn - 1;

                    // Scan each cell within the used range
                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        for (int col = firstColumn; col <= lastColumn; col++)
                        {
                            Cell cell = ws.Cells[row, col];
                            // Retrieve the cell's style
                            Style style = cell.GetStyle();

                            // Check if the font's theme color is Hyperlink (using string comparison for compatibility)
                            if (style.Font.ThemeColor.ToString() == "Hyperlink")
                            {
                                // Record worksheet name and cell address in the report
                                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                                reportSheet.Cells[reportRow, 1].PutValue(cell.Name); // e.g., "A1"
                                reportRow++;
                            }
                        }
                    }
                }

                // Auto‑fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the workbook (macro‑free)
                sourceWorkbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
