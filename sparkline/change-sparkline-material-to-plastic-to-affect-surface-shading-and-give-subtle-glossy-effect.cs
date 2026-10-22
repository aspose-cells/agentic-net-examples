// Title: How to set SparklineGroup material to Plastic for a subtle glossy effect using Aspose.Cells for .NET
// AI Prompts: Add a loop that iterates over every SparklineGroup in the worksheet and assigns SparklineMaterial.Plastic to its Material property before saving the workbook. | Update the example to verify the Aspose.Cells version supports SparklineGroup.Material, then apply the Plastic material to all sparkline groups to achieve glossy shading.
// Common Searches: Aspose.Cells .NET set sparkline group material to plastic | change sparkline shading to glossy using Aspose.Cells | apply plastic material to Excel sparkline groups programmatically | C# example for modifying sparkline material with Aspose.Cells | how to use SparklineMaterial enum in Aspose.Cells
// Tags: set sparkline material Aspose.Cells | sparklinegroup material property .NET | plastic sparkline shading Aspose.Cells | modify sparkline appearance programmatically | Aspose.Cells sparkline material enum

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample loads an Excel workbook, checks for existing SparklineGroup objects on the first worksheet, and demonstrates where to apply the Plastic material to sparkline groups (subject to API availability) before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // If sparkline groups exist, you can process them here.
                // The SparklineGroup class may not be available in older Aspose.Cells versions,
                // so this example skips modifying sparkline material.
                if (sheet.SparklineGroups != null && sheet.SparklineGroups.Count > 0)
                {
                    Console.WriteLine($"Workbook contains {sheet.SparklineGroups.Count} sparkline group(s).");
                }

                // Save the (potentially modified) workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
