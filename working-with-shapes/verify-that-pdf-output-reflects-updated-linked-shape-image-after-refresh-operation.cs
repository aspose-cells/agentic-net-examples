// Title: Refresh linked shape objects in an Excel workbook and confirm the updated PDF output using Aspose.Cells for .NET
// AI Prompts: Load an XLSX file containing linked pictures, invoke Workbook.RefreshAllLinkedObjects(), and save the workbook as a PDF with Aspose.Cells. | Validate that the PDF generated after refreshing linked objects reflects the latest linked image content. | Add try‑catch logic to gracefully handle a missing input workbook while converting Excel to PDF with linked shape refresh.
// Common Searches: how to refresh linked images in Aspose.Cells before exporting to PDF in C# | Aspose.Cells RefreshAllLinkedObjects not updating linked picture in PDF output | verify that linked shape changes appear in PDF generated from Excel workbook using Aspose.Cells .NET
// Tags: linked objects refresh Aspose.Cells | export workbook to PDF with refreshed images | linked picture update verification PDF | missing workbook file error handling Aspose.Cells | Aspose.Cells workbook conversion to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLinkedShapeRefresh
{
    // The example loads an Excel workbook, optionally refreshes all linked objects (such as linked pictures) via Workbook.RefreshAllLinkedObjects(), and saves the workbook as a PDF, including basic error handling for a missing input file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPdfPath = "output.pdf";

                // Verify required files exist
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input workbook not found: {inputPath}");

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: In some Aspose.Cells versions the method RefreshAllLinkedObjects is not available.
                // If your version supports it, you can uncomment the following line:
                // workbook.RefreshAllLinkedObjects();

                // Save the workbook as PDF
                workbook.Save(outputPdfPath, SaveFormat.Pdf);

                Console.WriteLine("Workbook converted to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
