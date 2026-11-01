// Title: Hide columns K‑M in an Excel worksheet and export to PDF with Aspose.Cells for .NET (handling hidden column visibility)
// AI Prompts: Generate C# code that uses Aspose.Cells to hide columns K through M in a workbook and then saves the worksheet as a PDF file. | Demonstrate how to set PdfSaveOptions in Aspose.Cells so that hidden columns are either excluded or included during Excel‑to‑PDF conversion.
// Common Searches: Aspose.Cells C# hide specific columns before converting Excel to PDF | How to keep hidden columns visible in PDF output using Aspose.Cells | C# example for hiding columns K-M and creating a PDF from an Excel file with Aspose.Cells
// Tags: hide columns Aspose.Cells C# | Excel to PDF conversion hidden columns Aspose.Cells | PdfSaveOptions HideHiddenColumns property | column visibility handling Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, hides columns K through M on the first worksheet with the HideColumns method, configures PdfSaveOptions, and saves the result as a PDF. By default hidden columns are omitted in the PDF, and the code notes how newer Aspose.Cells versions can control this behavior via HideHiddenColumns settings.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns K (index 10) through M (index 12) – total 3 columns
            sheet.Cells.HideColumns(10, 3);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Note: In this version of Aspose.Cells, hidden columns/rows are omitted by default.
            // If a newer version provides HideHiddenColumns/HideHiddenRows properties, they can be set here.

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
