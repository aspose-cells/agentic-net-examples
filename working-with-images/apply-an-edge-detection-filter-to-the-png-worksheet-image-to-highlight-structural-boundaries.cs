// Title: Export the first worksheet of an Excel workbook to a PNG image with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save first worksheet as PNG image | How to convert Excel sheet to PNG using Aspose.Cells .NET | Aspose.Cells export worksheet to PNG file programmatically | Save Excel worksheet as high‑resolution PNG in C# | Aspose.Cells example for generating PNG from workbook
// Tags: export worksheet to PNG Aspose.Cells | C# convert Excel sheet to image | Aspose.Cells save first worksheet as PNG | generate PNG from workbook .NET | Excel to PNG conversion using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace EdgeDetectionExample
{
    // The sample checks for the presence of input.xlsx, loads it with Aspose.Cells, and saves the first worksheet as a PNG image named sheet.png, handling any runtime exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "sheet.png";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Save the first worksheet as a PNG image
                workbook.Save(outputPath, SaveFormat.Png);

                Console.WriteLine($"Sheet image saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
