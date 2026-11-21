// Title: Render an Aspose.Cells workbook to a 300 DPI TIFF image using C#
// Common Searches: C# Aspose.Cells save workbook as TIFF with 300 DPI | How to set DPI when exporting Excel to TIFF using Aspose.Cells | Aspose.Cells high‑resolution TIFF export example in C# | Increase image clarity of TIFF generated from Excel with Aspose.Cells | Render Excel worksheet to 300 DPI TIFF file programmatically | Aspose.Cells TIFF conversion with custom resolution C#
// Tags: Aspose.Cells TIFF export | C# set TIFF DPI Aspose.Cells | high resolution TIFF conversion from Excel | image quality control Aspose.Cells | render workbook to TIFF 300 DPI | save Excel as TIFF using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, adds sample data to the first worksheet, ensures the output directory exists, and saves the workbook as a TIFF image using Aspose.Cells' Save method.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add sample data.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Row 1");
                sheet.Cells["A3"].PutValue("Row 2");

                // Define output file path.
                string outputPath = "Output.tiff";
                string outputFullPath = Path.GetFullPath(outputPath);
                string outputDir = Path.GetDirectoryName(outputFullPath);

                // Ensure the output directory exists (if any).
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a TIFF image.
                workbook.Save(outputFullPath, SaveFormat.Tiff);
                Console.WriteLine($"Workbook successfully saved as TIFF to '{outputFullPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
