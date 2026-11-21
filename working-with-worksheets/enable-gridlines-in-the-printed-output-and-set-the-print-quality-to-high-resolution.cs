// Title: How to enable printed gridlines for a worksheet and configure high‑resolution printing with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable gridlines in printed worksheet | set high resolution print options using Aspose.Cells .NET | page setup print gridlines and high quality output Aspose.Cells | C# Aspose.Cells print worksheet with gridlines and high DPI | configure printer settings for high resolution when saving Excel with Aspose.Cells
// Tags: Aspose.Cells page setup print gridlines | C# Aspose.Cells high resolution printing | worksheet print options Aspose.Cells | enable gridlines in printed Excel Aspose.Cells | Aspose.Cells printer settings high DPI

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses the first Worksheet, turns on PageSetup.PrintGridlines to show gridlines in the printed output, notes that direct print‑quality control is not available in recent Aspose.Cells versions and must be handled via external printer settings, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Enable gridlines in the printed output for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintGridlines = true;

            // NOTE: The PrintQuality property is not available in recent Aspose.Cells versions.
            // If needed, configure printer settings externally.

            // Define output file path
            string outputPath = "PrintedWithGridlines.xlsx";

            try
            {
                // Save the workbook (lifecycle: save)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception e)
        {
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }
}
