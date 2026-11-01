// Title: C# program to load an Excel workbook with Italian locale, add a subtotal row, and export it to PDF using Aspose.Cells
// Common Searches: how to load an Excel workbook with Italian locale using Aspose.Cells in C# | add a total row to a worksheet and convert it to PDF with Aspose.Cells | set LoadOptions CultureInfo to it-IT when processing Excel files in .NET | calculate sum of a column and append a subtotal before saving as PDF using Aspose.Cells | export localized Excel data to PDF with Aspose.Cells for .NET
// Tags: loadoptions italian culture aspnet cells | add subtotal row aspnet cells | excel to pdf conversion with localization aspnet cells | column sum calculation aspnet cells | export workbook pdf aspnet cells

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example checks for an input.xlsx file, creates a simple workbook if missing, loads it with LoadOptions specifying the Italian culture (it-IT), computes the sum of the 'Amount' column, inserts a 'Total' row, and saves the modified workbook as a PDF (output.pdf) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input file exists; create a simple workbook if it does not
            if (!File.Exists(inputPath))
            {
                var sampleWb = new Workbook();
                var ws = sampleWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Amount");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("A");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("B");
                ws.Cells["B4"].PutValue(30);
                sampleWb.Save(inputPath);
            }

            // Set the Italian culture for loading the workbook
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = new CultureInfo("it-IT")
            };

            // Load the workbook
            var workbook = new Workbook(inputPath, loadOptions);
            var worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxColumn = worksheet.Cells.MaxDataColumn;

            // ----- Subtotal logic (manual implementation) -----
            // Compute sum of the "Amount" column (index 1) for all data rows
            double total = 0;
            for (int row = 1; row <= maxRow; row++) // start from 1 to skip header
            {
                var cell = worksheet.Cells[row, 1];
                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                {
                    total += val;
                }
            }

            // Add a total row after the last data row
            int totalRowIndex = maxRow + 1;
            worksheet.Cells[totalRowIndex, 0].PutValue("Total");
            worksheet.Cells[totalRowIndex, 1].PutValue(total);
            // --------------------------------------------------

            // Save the result as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
