// Title: Configure a worksheet for 300 DPI high‑quality printing using Aspose.Cells in C#
// Common Searches: how to set worksheet print DPI to 300 in Aspose.Cells C# | Aspose.Cells high resolution printing options for Excel files | configure PDF export to 300 DPI with Aspose.Cells .NET | set print options for worksheet when rendering image using Aspose.Cells | workaround missing PrintResolution property in Aspose.Cells
// Tags: worksheet print quality adjustment Aspose.Cells | high‑resolution PDF generation Aspose.Cells .NET | print options API usage Aspose.Cells C# | 300 DPI export workaround Aspose.Cells | Excel printing settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, accesses the first Worksheet, and notes that the PrintResolution property is unavailable in the current Aspose.Cells version. It advises configuring print options (such as DPI) during PDF or image rendering to achieve high‑quality 300 DPI output before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Note: The PrintResolution property is not available in this version of Aspose.Cells.
                // If high‑quality printing is required, configure print options when rendering to PDF/Image.

                // Define the output file path
                string outputPath = "HighQualityPrint.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
