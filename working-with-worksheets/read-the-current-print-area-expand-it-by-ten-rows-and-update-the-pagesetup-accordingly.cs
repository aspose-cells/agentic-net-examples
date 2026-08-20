// Title: Expand Excel Print Area by 10 Rows with Aspose.Cells for .NET
// Description: Loads a workbook, reads the first worksheet's PageSetup.PrintArea, creates a Range from that address, extends the bottom row by ten rows, updates the PrintArea string, and saves the file.
// Keywords: Aspose.Cells C# print area | expand Excel print range .NET | PageSetup.PrintArea modify | add rows to Excel print area | Aspose.Cells range manipulation
// Common Searches: how to increase Excel print area using Aspose.Cells | C# expand worksheet print range by rows | Aspose.Cells set new PrintArea programmatically | extend Excel print area ten rows C#
// Developer Intent: Read the current print area of a worksheet, enlarge it by ten rows, and save the workbook with the updated PageSetup.
// Use Cases: Append blank rows for notes in a report template before printing. | Automatically grow the print range when new data rows are added during batch report generation. | Maintain consistent page breaks after dynamically extending the printable region.
// AI Prompts: Generate C# code using Aspose.Cells that reads a worksheet’s PrintArea, expands it by a configurable number of rows, and writes the new range back. | Create a reusable method that accepts a workbook path, sheet index, and row offset, then returns the updated PrintArea string. | Provide robust error handling for scenarios where a worksheet has no PrintArea defined before attempting to expand it.

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, reads the first worksheet's PageSetup.PrintArea, creates a Range from that address, extends the bottom row by ten rows, updates the PrintArea string, and saves the file.
class ExpandPrintArea
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the current print area
            string currentPrintArea = worksheet.PageSetup.PrintArea;

            if (string.IsNullOrEmpty(currentPrintArea))
            {
                Console.WriteLine("Print area is not set.");
                return;
            }

            // Create a range object from the current print area (fully qualified to avoid ambiguity)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange(currentPrintArea);

            int startRow = range.FirstRow;          // zero‑based
            int startColumn = range.FirstColumn;    // zero‑based
            int rowCount = range.RowCount;
            int columnCount = range.ColumnCount;

            // Expand the end row by ten rows
            int newEndRow = startRow + rowCount - 1 + 10; // zero‑based index of the new last row

            // Build the new print area string (e.g., "A1:B20")
            string startCell = CellsHelper.ColumnIndexToName(startColumn) + (startRow + 1);
            string endCell = CellsHelper.ColumnIndexToName(startColumn + columnCount - 1) + (newEndRow + 1);
            string newPrintArea = $"{startCell}:{endCell}";

            // Update the worksheet's PageSetup with the expanded print area
            worksheet.PageSetup.PrintArea = newPrintArea;

            // Save the workbook
            workbook.Save(outputPath);

            Console.WriteLine($"Print area expanded from \"{currentPrintArea}\" to \"{newPrintArea}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
