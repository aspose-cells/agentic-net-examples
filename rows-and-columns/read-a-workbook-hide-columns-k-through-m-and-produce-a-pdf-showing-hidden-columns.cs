// Title: Hide columns K to M in an Excel worksheet and export to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide specific columns and export to PDF | How to render hidden columns in PDF using Aspose.Cells .NET | C# code to hide columns K-M in Excel and generate PDF with Aspose.Cells | Export Excel worksheet to PDF while keeping hidden columns visible with Aspose.Cells | PdfSaveOptions RenderHiddenColumns property usage in Aspose.Cells C#
// Tags: hide column range Aspose.Cells | export worksheet to PDF Aspose.Cells | PdfSaveOptions render hidden columns | C# hide Excel columns before PDF conversion | Aspose.Cells column visibility handling

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, hides columns K through M on the first worksheet using Aspose.Cells, configures PDF save options, and saves the workbook as a PDF while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            var sheet = workbook.Worksheets[0];

            // Hide columns K (index 10) through M (index 12) inclusive
            for (int col = 10; col <= 12; col++)
            {
                sheet.Cells.HideColumn(col);
            }

            // Configure PDF save options
            var pdfOptions = new PdfSaveOptions
            {
                // In older Aspose.Cells versions RenderHiddenColumns/Rows may not be available.
                // If needed, these properties can be set when supported.
                // RenderHiddenColumns = true,
                // RenderHiddenRows = true
            };

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
