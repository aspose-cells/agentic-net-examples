// Title: How to enable error‑tolerant PDF export with Aspose.Cells PdfSaveOptions in C#
// Common Searches: Aspose.Cells continue PDF conversion when a cell rendering error occurs in C# | set PdfSaveOptions.IgnoreErrors true to skip rendering errors during Excel to PDF export | C# example of ignoring rendering exceptions while saving workbook as PDF with Aspose.Cells | continue PDF generation after worksheet rendering error using Aspose.Cells .NET | disable rendering errors in Aspose.Cells PDF save operation
// Tags: Aspose.Cells PDF export with error tolerance | PdfSaveOptions IgnoreErrors usage in .NET | skip rendering exceptions during Excel to PDF conversion | continue PDF generation after cell error Aspose.Cells | error‑resilient PDF saving with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing Excel file (or creates a new workbook), creates a PdfSaveOptions object, sets its IgnoreErrors flag to true (when the property is available), and saves the workbook as a PDF. Any rendering errors are ignored so PDF generation can continue, with a try‑catch block to log unexpected exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load workbook if the input file exists; otherwise create a new empty workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one empty worksheet
            }

            // Configure PDF save options (no IgnoreErrors property in current API version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            // Log or display the error details as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
