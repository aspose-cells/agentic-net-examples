// Title: C# – Apply 95% Zoom to Worksheets with More Than 500 Rows and Export to PDF with Aspose.Cells
// Description: Loads an Excel workbook, iterates each worksheet, sets the view zoom to 95% when the sheet contains over 500 rows, and saves the entire workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set worksheet zoom | conditional zoom Excel rows | export Excel to PDF Aspose.Cells | iterate worksheets .NET | zoom 95% before PDF conversion
// Common Searches: Aspose.Cells set zoom for large worksheets | C# export Excel to PDF after adjusting zoom | how to apply conditional zoom in Aspose.Cells | loop through worksheets and change view settings | save workbook as PDF with Aspose.Cells .NET
// Developer Intent: Set a 95% view zoom on any worksheet that exceeds 500 rows and then generate a PDF of the workbook.
// Use Cases: Create printable PDFs where dense sheets are automatically zoomed to fit more data per page. | Automate batch reporting that adjusts worksheet zoom based on row count before PDF generation. | Build a server‑side service that validates incoming Excel files, applies conditional zoom, and returns a PDF version.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets, sets Zoom = 95 for sheets with more than 500 rows, and saves the workbook as a PDF. | Provide an example that includes error handling for missing input files while applying conditional zoom and exporting to PDF. | Show how to customize PDF save options (orientation, compression, page size) after setting worksheet zoom with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, iterates each worksheet, sets the view zoom to 95% when the sheet contains over 500 rows, and saves the entire workbook as a PDF using Aspose.Cells for .NET.
class ApplyZoomAndExportPdf
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and set zoom if rows > 500
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int usedRows = sheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based
                if (usedRows > 500)
                {
                    sheet.Zoom = 95; // Set view zoom to 95%
                }
            }

            // Save the workbook as a PDF with default options
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine("PDF exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
