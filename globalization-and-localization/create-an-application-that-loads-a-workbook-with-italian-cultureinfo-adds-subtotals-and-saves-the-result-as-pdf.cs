// Title: Load an Excel workbook with Italian locale, add subtotal SUM formulas, and export to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook with it-IT CultureInfo | Insert subtotal row with SUM formulas in an Aspose.Cells worksheet | Export Excel to PDF after adding subtotals using Aspose.Cells .NET | Set CultureInfo in LoadOptions when opening an Excel file with Aspose.Cells | Programmatically create subtotal calculations in C# with Aspose.Cells
// Tags: cultureinfo loadoptions workbook Aspose.Cells | subtotal sum formula insertion Aspose.Cells | excel to pdf conversion Aspose.Cells C# | italian locale handling Aspose.Cells | c# insert subtotal calculations worksheet

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an input XLSX file, loads it with LoadOptions configured for the Italian (it-IT) culture, adds SUM formulas for selected columns in a subtotal row, labels the row, and saves the workbook as a PDF file using Aspose.Cells for .NET, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            const string inputPath = "input.xlsx";
            const string outputPdf = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            // Set Italian culture for loading the workbook
            CultureInfo italianCulture = new CultureInfo("it-IT");
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = italianCulture
            };

            // Load the workbook
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for which subtotals will be calculated
            // Example: rows 0‑9 (first 10 rows) and columns 0‑3 (first 4 columns)
            int startRow = 0;
            int startColumn = 0;
            int totalRows = 9;      // zero‑based index of the last row in the range
            int totalColumns = 3;   // zero‑based index of the last column in the range

            // Columns (zero‑based) for which we want to calculate sums
            int[] subtotalColumns = new int[] { 1, 2 };

            // Row where the subtotal formulas will be placed (just after the data range)
            int subtotalRow = totalRows + 1;

            // Insert SUM formulas for each specified column
            foreach (int col in subtotalColumns)
            {
                // Build cell references for the start and end of the range (e.g., A1:A10)
                string startCell = CellsHelper.CellIndexToName(startRow, col);
                string endCell = CellsHelper.CellIndexToName(totalRows, col);

                // Set the formula in the subtotal row
                sheet.Cells[subtotalRow, col].Formula = $"=SUM({startCell}:{endCell})";
            }

            // Optionally, add a label for the subtotal row
            sheet.Cells[subtotalRow, startColumn].PutValue("Subtotal");

            // Save the workbook as a PDF file
            workbook.Save(outputPdf, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPdf}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
