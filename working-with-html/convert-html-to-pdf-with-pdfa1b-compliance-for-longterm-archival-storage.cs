// Title: Convert HTML to PDF/A‑1b with Aspose.Cells for .NET (C#) – archival‑ready PDF generation
// Common Searches: C# Aspose.Cells convert HTML file to PDF/A-1b compliant document | How to generate PDF/A-1b from HTML using Aspose.Cells in .NET | Aspose.Cells PDF/A-1b export options for HTML workbook C# | Create archival PDF/A-1b from HTML with one page per worksheet Aspose.Cells | Embedding Windows fonts for PDF/A-1b conversion with Aspose.Cells C#
// Tags: Aspose.Cells HTML to PDF/A-1b conversion | PdfSaveOptions Compliance PdfA1b C# | LoadOptions LoadFormat Html Aspose.Cells | EmbedStandardWindowsFonts PDF/A Aspose.Cells | OnePagePerSheet PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfA1b
{
    // The example loads an HTML file into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Html, configures PdfSaveOptions to enforce PDF/A‑1b compliance (including embedding standard Windows fonts and one page per worksheet), and saves the result as an archival‑ready PDF/A‑1b file. Input and output paths can be supplied via command‑line arguments or default to files in the executable directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input HTML file and output PDF file.
            string inputHtmlPath;
            string outputPdfPath;

            if (args.Length >= 2)
            {
                inputHtmlPath = args[0];
                outputPdfPath = args[1];
            }
            else
            {
                // Default file names in the executable directory.
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                inputHtmlPath = Path.Combine(exeDir, "input.html");
                outputPdfPath = Path.Combine(exeDir, "output.pdf");
                Console.WriteLine("Usage: HtmlToPdfA1b <inputHtmlPath> <outputPdfPath>");
                Console.WriteLine("No arguments supplied – using default files:");
                Console.WriteLine($"  Input HTML : {inputHtmlPath}");
                Console.WriteLine($"  Output PDF : {outputPdfPath}");
            }

            // Validate input file.
            if (!File.Exists(inputHtmlPath))
            {
                Console.Error.WriteLine($"Error: Input HTML file not found: {inputHtmlPath}");
                return;
            }

            try
            {
                // Load the HTML file into a workbook.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

                // Configure PDF save options for PDF/A‑1b compliance.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // PDF/A‑1b compliance ensures long‑term archival suitability.
                    Compliance = PdfCompliance.PdfA1b,

                    // Optional: generate one PDF page per worksheet.
                    OnePagePerSheet = true,

                    // Optional: embed standard Windows fonts (required for PDF/A).
                    EmbedStandardWindowsFonts = true
                };

                // Save the workbook as PDF/A‑1b.
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"Successfully converted HTML to PDF/A‑1b:");
                Console.WriteLine($"  Source : {inputHtmlPath}");
                Console.WriteLine($"  Target : {outputPdfPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during conversion:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
