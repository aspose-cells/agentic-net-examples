using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfWithPrintJs
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve document structure for accessibility
                    ExportDocumentStructure = true
                };

                // Embed JavaScript for auto‑print if the property is available in the current Aspose.Cells version
                var jsProperty = typeof(PdfSaveOptions).GetProperty("Javascript");
                if (jsProperty != null && jsProperty.CanWrite)
                {
                    jsProperty.SetValue(pdfOptions, "this.print({bUI:false,bSilent:true});");
                }
                else
                {
                    Console.WriteLine("Warning: PdfSaveOptions does not support JavaScript embedding in this Aspose.Cells version.");
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully converted to PDF with auto‑print JavaScript (if supported). Output: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}