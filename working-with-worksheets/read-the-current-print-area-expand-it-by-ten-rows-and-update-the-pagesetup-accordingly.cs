using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPrintAreaExpand
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Sample data and initial print area (A1:C5)
                // ------------------------------------------------------------
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }
                worksheet.PageSetup.PrintArea = "A1:C5";

                // ------------------------------------------------------------
                // 1. Read the current print area
                // ------------------------------------------------------------
                string currentPrintArea = worksheet.PageSetup.PrintArea; // e.g. "A1:C5"
                Console.WriteLine($"Current PrintArea: {currentPrintArea}");

                // ------------------------------------------------------------
                // 2. Create a Range object from the print area to obtain its bounds
                // ------------------------------------------------------------
                AsposeRange range = worksheet.Cells.CreateRange(currentPrintArea);

                // Starting row/column (zero‑based)
                int startRow = range.FirstRow;
                int startColumn = range.FirstColumn;

                // Ending row/column (zero‑based)
                int endRow = startRow + range.RowCount - 1;
                int endColumn = startColumn + range.ColumnCount - 1;

                // ------------------------------------------------------------
                // 3. Expand the area by ten rows (keep the same columns)
                // ------------------------------------------------------------
                int expandedEndRow = endRow + 10; // add ten rows

                // Convert column indexes back to column letters
                string startColLetter = CellsHelper.ColumnIndexToName(startColumn);
                string endColLetter = CellsHelper.ColumnIndexToName(endColumn);

                // Build the new print area string (Excel uses 1‑based row numbers)
                string newPrintArea = $"{startColLetter}{startRow + 1}:{endColLetter}{expandedEndRow + 1}";
                Console.WriteLine($"Expanded PrintArea: {newPrintArea}");

                // ------------------------------------------------------------
                // 4. Update the worksheet PageSetup with the new print area
                // ------------------------------------------------------------
                worksheet.PageSetup.PrintArea = newPrintArea;

                // ------------------------------------------------------------
                // 5. Save the workbook (demonstrates that the new print area is persisted)
                // ------------------------------------------------------------
                string outputPath = "PrintAreaExpanded.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with expanded print area at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}