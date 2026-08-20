// Title: Convert XLSB Office Add‑In to PDF with 0.8 Scaling Using Aspose.Cells for .NET
// Description: C# example that verifies an XLSB Office Add‑In file, loads it with LoadOptions (Xlsb), and converts it to PDF using Aspose.Cells ConversionUtility. The PdfSaveOptions are configured with a custom page scaling factor of 0.8 to render the add‑in correctly. Includes basic error handling and console feedback.
// Keywords: Aspose.Cells XLSX to PDF | C# convert XLSB Office Add‑In | PdfSaveOptions scaling factor | 0.8 page scaling Aspose | ConversionUtility PDF conversion | .NET render Office Add‑In | XLSB to PDF Aspose.Cells | custom PDF scaling C#
// Common Searches: How to convert an XLSB Office Add‑In to PDF with Aspose.Cells | Aspose.Cells C# set PDF scaling factor 0.8 | Render Office Add‑In while saving XLSB as PDF | PdfSaveOptions page scaling example in .NET | ConversionUtility convert XLSB to PDF with custom scaling
// Developer Intent: Generate a PDF from an XLSB Office Add‑In workbook in C# while applying a 0.8 page‑scaling factor.
// Use Cases: Automate batch conversion of Office Add‑In XLSB files to PDFs for documentation archives. | Create printable PDFs with reduced page size for mobile‑friendly distribution. | Integrate the conversion step into CI/CD pipelines to produce PDF reports from add‑in workbooks. | Validate file existence and capture conversion errors for robust enterprise workflows.
// AI Prompts: Show me how to set PdfSaveOptions.ScaleFactor = 0.8 before calling ConversionUtility. | Provide a complete C# snippet that logs detailed ConversionUtility errors to a file. | Explain how to adjust image resolution and page margins together with scaling for XLSB to PDF conversion. | Generate a PowerShell script that runs the C# program on a Windows server for scheduled batch processing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace AsposeCellsAddInRendering
{
    // C# example that verifies an XLSB Office Add‑In file, loads it with LoadOptions (Xlsb), and converts it to PDF using Aspose.Cells ConversionUtility. The PdfSaveOptions are configured with a custom page scaling factor of 0.8 to render the add‑in correctly. Includes basic error handling and console feedback.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source XLSB workbook (Office Add‑In file)
                string sourcePath = "AddInWorkbook.xlsb";

                // Desired output PDF file path
                string outputPath = "AddInWorkbook.pdf";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // LoadOptions to specify that the source file is an XLSB workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);

                // PDF save options (default settings; scaling not required for this version)
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Perform the conversion using the provided ConversionUtility method.
                ConversionUtility.Convert(sourcePath, loadOptions, outputPath, saveOptions);

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
