// Title: Generate an Excel report of cells with QuotePrefix style set to true using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, identifies cells where Style.QuotePrefix is true, and writes their zero‑based row, column, and address to a new worksheet named "QuotePrefixReport". | Extend the program to also record the source worksheet name for each quoted‑prefix cell in the generated report. | Adjust the output so that row and column numbers are 1‑based while keeping the cell address unchanged.
// Common Searches: asp.net how to export list of cells with QuotePrefix flag using Aspose.Cells | c# Aspose.Cells find all cells with quoted prefix style in a workbook | generate report of quoted prefix cells with row and column indices in Excel using Aspose | Aspose.Cells enumerate cells where Style.QuotePrefix is true across worksheets
// Tags: Aspose.Cells enumerate QuotePrefix cells | C# create Excel summary of quoted‑prefix style | Aspose.Cells write row column address to report sheet | Excel workbook scan for QuotePrefix attribute | Aspose.Cells generate styled‑cell report

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an input workbook, iterates through each worksheet (excluding the report sheet), checks every cell's Style.QuotePrefix flag, and records the zero‑based row index, column index, and cell address of matching cells on a new worksheet called "QuotePrefixReport" before saving the workbook.
class QuotePrefixReportGenerator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "QuotePrefixReport";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Row Index");
            reportSheet.Cells[0, 1].PutValue("Column Index");
            reportSheet.Cells[0, 2].PutValue("Cell Address");

            int reportRow = 1; // start after header

            // Iterate through all worksheets in the source workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the report sheet itself
                if (sheet.Name == reportSheet.Name) continue;

                // Get the used range of the worksheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Retrieve the cell's style
                    Style style = cell.GetStyle();

                    // Check if QuotePrefix is set to true
                    if (style.QuotePrefix)
                    {
                        // Record row and column indices (zero‑based)
                        reportSheet.Cells[reportRow, 0].PutValue(cell.Row);
                        reportSheet.Cells[reportRow, 1].PutValue(cell.Column);
                        reportSheet.Cells[reportRow, 2].PutValue(cell.Name); // e.g., "A1"

                        reportRow++;
                    }
                }
            }

            // Save the workbook with the report
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
