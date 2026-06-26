using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace MyApp
{
    class ValidateTitleBeforeExport
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Ensure the Title property is set (required for PDF export)
                if (string.IsNullOrWhiteSpace(workbook.BuiltInDocumentProperties.Title))
                {
                    workbook.BuiltInDocumentProperties.Title = "Untitled Document";
                }

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DisplayDocTitle = true
                };

                // Save the workbook as PDF
                string outputPath = "ExportedDocument.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}