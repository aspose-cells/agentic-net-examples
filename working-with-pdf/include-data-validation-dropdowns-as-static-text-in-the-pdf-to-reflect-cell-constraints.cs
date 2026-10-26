// Title: Convert an Excel workbook to PDF with Aspose.Cells for .NET using PdfSaveOptions
// Common Searches: how to save an Excel file as PDF in C# with Aspose.Cells | Aspose.Cells C# export workbook to PDF with each sheet on a separate page | render Excel data validation list as plain text in PDF using Aspose.Cells | C# PdfSaveOptions OnePagePerSheet example with Aspose.Cells | convert Excel to PDF while preserving cell formatting in .NET
// Tags: aspocells excel to pdf conversion | pdfsaveoptions onepagepersheet c# | static rendering of data validation lists pdf | c# export workbook as pdf aspocells | preserve cell formatting during pdf export aspocells | aspocells pdf export with validation constraints

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook (or creates a simple one), configures PdfSaveOptions to place each worksheet on a single page, and saves the workbook as a PDF using Aspose.Cells for .NET. It illustrates the core steps for Excel‑to‑PDF conversion, which can be extended to render data‑validation dropdowns as static text in the PDF.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.pdf";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a simple workbook.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sample Data");
                }

                // Configure PDF save options.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example option: render each worksheet on a separate page.
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF using the configured options.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
