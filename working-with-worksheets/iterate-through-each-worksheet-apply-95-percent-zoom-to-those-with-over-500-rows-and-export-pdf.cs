// Title: C# – Apply 95% zoom to worksheets with more than 500 rows and save the workbook as PDF using Aspose.Cells
// Common Searches: how to set worksheet zoom to 95% for large sheets with Aspose.Cells in C# | export Excel workbook to PDF after adjusting zoom for worksheets that exceed 500 rows in .NET | iterate through all worksheets and change view zoom based on row count using Aspose.Cells | conditional worksheet zoom before PDF conversion with Aspose.Cells for C#
// Tags: worksheet zoom adjustment Aspose.Cells C# | row‑count based zoom Aspose.Cells | PDF export workbook Aspose.Cells | iterate worksheets modify view Aspose.Cells | large sheet zoom optimization Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file, loops through each worksheet, checks the total number of rows, sets the worksheet zoom to 95 % when the row count exceeds 500, and then saves the entire workbook as a PDF, with basic file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the existing file
            var workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow is zero‑based; add 1 to get the total row count
                int rowCount = sheet.Cells.MaxDataRow + 1;

                // Apply 95% zoom if the sheet contains more than 500 rows
                if (rowCount > 500)
                {
                    sheet.Zoom = 95; // Zoom expects an integer percentage
                }
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
