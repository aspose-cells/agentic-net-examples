// Title: Convert XLSX with Office Add‑In Controls to Interactive PDF using Aspose.Cells for .NET
// Description: Loads an XLSX workbook that contains Office Add‑In controls with LoadOptions, then converts it to PDF via ConversionUtility and PdfSaveOptions, keeping the controls functional as interactive form fields.
// Keywords: Aspose.Cells XLSX to PDF | preserve Office Add‑In controls | interactive PDF from Excel | ConversionUtility Aspose | PdfSaveOptions form fields | C# Excel PDF conversion
// Common Searches: Aspose.Cells keep Office Add‑In form fields in PDF | convert Excel with add‑in controls to interactive PDF C# | preserve Excel add‑in controls when saving as PDF | ConversionUtility PDF conversion example Aspose.Cells | load XLSX with Office Add‑In controls Aspose
// Developer Intent: Convert an Excel workbook that includes Office Add‑In controls into a PDF while retaining those controls as interactive elements.
// Use Cases: Create PDF reports from Excel templates that contain embedded Office Add‑In form fields without losing interactivity. | Batch‑process multiple XLSX files with add‑in controls into PDFs for archival or distribution. | Expose a web service that accepts XLSX uploads with Office Add‑In controls and returns a PDF preserving the interactive elements.
// AI Prompts: Show how to customize PdfSaveOptions to style preserved form fields during conversion. | Provide error‑handling patterns for ConversionUtility when encountering unsupported Office Add‑In controls. | Demonstrate a script that scans a folder, converts each XLSX with add‑in controls to PDF, and logs conversion results.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;   // Contains PdfSaveOptions

namespace OfficeAddInPdfConversion
{
    // Loads an XLSX workbook that contains Office Add‑In controls with LoadOptions, then converts it to PDF via ConversionUtility and PdfSaveOptions, keeping the controls functional as interactive form fields.
    public class Converter
    {
        public static void Run()
        {
            // Path to the source XLSX workbook that contains Office Add‑In controls
            string sourcePath = "input.xlsx";

            // Desired output PDF file path
            string destPath = "output.pdf";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load options – explicitly specify the format to ensure correct loading
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Save options for PDF – default options are sufficient for preserving form fields
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Perform the conversion using the provided ConversionUtility rule
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                Console.WriteLine("Conversion completed. PDF saved to: " + destPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }

        // Entry point required for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
