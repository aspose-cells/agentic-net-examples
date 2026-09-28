// Title: Embed a text file as an attachment in a PDF generated from an Aspose.Cells workbook using C# PdfSaveOptions
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, adds sample data, and saves it as a PDF with a .txt file attached via PdfSaveOptions. | Describe the steps to use the PdfSaveOptions.Attachments collection to embed an external text document into the PDF produced by Aspose.Cells.
// Common Searches: how to add a .txt attachment to a PDF created with Aspose.Cells in C# | Aspose.Cells PdfSaveOptions embed external file into PDF example | C# generate PDF from workbook with attached text document using Aspose | attach files to PDF when exporting Excel with Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions embed attachment | C# generate PDF with attached text file | Aspose.Cells export workbook to PDF with attachments | PdfSaveOptions.Attachments usage .NET | embed external document in PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to create an Aspose.Cells Workbook, populate it with data, configure PdfSaveOptions to include a text file as an embedded attachment, and save the workbook as a PDF. It also shows directory preparation and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            const string outputPdf = "Result.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPdf));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPdf, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPdf}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
