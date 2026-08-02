// Title: C# – Convert XLSX with Office Add‑Ins to PDF using Aspose.Cells default scaling
// Description: Loads an XLSX workbook that may contain Office Add‑Ins (charts, embedded objects, etc.) and converts it to PDF with Aspose.Cells.Utility.ConversionUtility.Convert, relying on the library's default page‑scaling settings to preserve the original layout.
// Keywords: Aspose.Cells C# PDF conversion | XLSX to PDF default scaling | render Office Add‑Ins | ConversionUtility Convert example | preserve embedded objects Excel PDF
// Common Searches: Aspose.Cells convert Excel with add‑ins to PDF C# | ConversionUtility default page scaling example | How to export XLSX charts to PDF using Aspose | C# code to render Office Add‑Ins in PDF conversion
// Developer Intent: Generate a PDF from an Excel file that includes Office Add‑Ins without specifying custom scaling.
// Use Cases: Create printable PDF reports from Excel templates that contain embedded charts or SmartArt. | Batch‑process multiple workbooks with add‑ins while keeping their original layout. | Expose a web service that receives an XLSX file and returns a PDF rendering of all embedded objects.
// AI Prompts: Provide C# code that uses Aspose.Cells ConversionUtility to convert an XLSX with Office Add‑Ins to PDF and captures conversion errors. | Explain how Aspose.Cells handles page scaling by default when exporting Excel to PDF and how to change it if needed. | Show a verification script that checks whether embedded objects from the source workbook appear in the generated PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInPdfConversion
{
    // Loads an XLSX workbook that may contain Office Add‑Ins (charts, embedded objects, etc.) and converts it to PDF with Aspose.Cells.Utility.ConversionUtility.Convert, relying on the library's default page‑scaling settings to preserve the original layout.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Path where the resulting PDF will be saved
            string pdfPath = "output.pdf";

            try
            {
                // Load the workbook (creates a Workbook instance from the XLSX file)
                Workbook workbook = new Workbook(sourcePath);

                // The workbook may contain Office Add‑Ins (e.g., embedded objects, charts, etc.).
                // ConversionUtility.Convert will render the workbook to PDF using the default
                // page scaling (no explicit scaling is applied, which matches the requirement).
                ConversionUtility.Convert(sourcePath, pdfPath);

                Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
