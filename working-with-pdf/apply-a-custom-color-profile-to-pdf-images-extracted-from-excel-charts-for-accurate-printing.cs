// Title: Convert an Excel workbook to PDF while preserving chart images with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook as PDF with embedded charts | How to export Excel charts to PDF using Aspose.Cells .NET | C# convert .xlsx to PDF preserving chart graphics Aspose | Save Excel file to PDF including chart objects with Aspose.Cells | Generate PDF from Excel workbook that keeps chart images in C#
// Tags: Aspose.Cells workbook to PDF conversion with charts | C# export Excel charts to PDF using Aspose | SaveFormat.Pdf for chart preservation in Aspose.Cells | PDF generation from Excel workbook .NET | Aspose.Cells chart image handling in PDF output

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF using SaveFormat.Pdf, which automatically retains any chart images, while handling errors and reporting the output location.
class Program
{
    static void Main()
    {
        try
        {
            // Paths – adjust as needed
            string excelPath = "input.xlsx";
            string outputPdfPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);

            // Save the workbook as PDF (charts are preserved automatically)
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to: {outputPdfPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
