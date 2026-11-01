// Title: Set the PDF Title metadata to the Excel workbook name when converting to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, extracts the file name without extension, assigns it to the workbook's BuiltInDocumentProperties.Title, and saves the workbook as a PDF. | Show how to configure Aspose.Cells PdfSaveOptions to retain built‑in document properties while exporting an Excel workbook to PDF in a .NET project. | Provide a robust C# example that verifies the source Excel file exists, sets the PDF title metadata, and implements exception handling for the conversion process.
// Common Searches: aspocells c# set pdf title from excel filename during conversion | how to preserve document properties when exporting Excel to PDF with Aspose.Cells | C# example for assigning BuiltInDocumentProperties.Title before saving as PDF | error handling for missing input.xlsx in Aspose.Cells PDF export
// Tags: aspocells pdf title builtindocumentproperties | excel to pdf conversion metadata c# | pdfsaveoptions preserve document properties aspocells | set pdf document title from workbook name aspocells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The program loads 'input.xlsx', extracts its filename without the extension, assigns that value to the workbook's BuiltInDocumentProperties.Title, and saves the workbook as 'output.pdf' using Aspose.Cells, with basic error handling for missing files.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input Excel file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the original Excel workbook
                var workbook = new Workbook(inputPath);

                // Extract the workbook name without extension to use as PDF title
                string pdfTitle = Path.GetFileNameWithoutExtension(inputPath);

                // Set the PDF title via built‑in document properties
                workbook.BuiltInDocumentProperties.Title = pdfTitle;

                // Configure PDF save options (default options are sufficient here)
                var pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified title metadata
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF file '{outputPath}' has been created successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
