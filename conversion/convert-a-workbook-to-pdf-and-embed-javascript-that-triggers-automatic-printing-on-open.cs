// Title: Convert an Excel workbook to PDF and embed auto‑print JavaScript using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to include a JavaScript snippet that calls this.print(true) when the PDF opens, and saves the result as a PDF. | Show how to use reflection in C# to detect the JavaScript (or Javascript) property on PdfSaveOptions and assign an auto‑print script before exporting the workbook to PDF.
// Common Searches: how to embed auto‑print JavaScript in PDF created from Excel with Aspose.Cells C# | Aspose.Cells PdfSaveOptions JavaScript property for printing on open | C# convert .xlsx to .pdf and add print script using Aspose.Cells | using reflection to set JavaScript on PdfSaveOptions in Aspose.Cells | check Aspose.Cells version support for JavaScript in PDF export
// Tags: Aspose.Cells PDF auto‑print JavaScript | PdfSaveOptions JavaScript property | C# Excel to PDF conversion with script | reflection set PdfSaveOptions property | Aspose.Cells PDF export embedded script

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfWithJavaScript
{
    // The example loads an input.xlsx workbook, creates PdfSaveOptions, uses reflection to locate the JavaScript (or Javascript) property, assigns the script "this.print(true);" for automatic printing, saves the workbook as output.pdf, and handles errors gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Attempt to embed JavaScript for auto‑print (property name may vary by version)
                var jsProperty = typeof(PdfSaveOptions).GetProperty("JavaScript") ??
                                 typeof(PdfSaveOptions).GetProperty("Javascript");

                if (jsProperty != null && jsProperty.CanWrite)
                {
                    jsProperty.SetValue(pdfOptions, "this.print(true);");
                }
                else
                {
                    Console.WriteLine("JavaScript embedding is not supported in this Aspose.Cells version.");
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook has been converted to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
