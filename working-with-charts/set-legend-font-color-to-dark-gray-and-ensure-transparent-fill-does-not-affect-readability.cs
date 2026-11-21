// Title: How to set a chart legend’s font color to dark gray and apply 50 % transparent fill with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change chart legend font color to dark gray | set transparent fill for chart legend using Aspose.Cells .NET | C# example for customizing Excel chart legend appearance with Aspose.Cells | apply 50% transparency to chart legend area in Aspose.Cells | modify legend fill type to solid and adjust transparency in Aspose.Cells chart
// Tags: chart legend font color Aspose.Cells | legend fill transparency Aspose.Cells .NET | solid fill format for chart legend Aspose.Cells | customize Excel chart legend C# | Aspose.Cells chart legend styling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;          // For Chart class
using Aspose.Cells.Drawing;        // For FillType enum
using System.Drawing;             // For Color

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, accesses the first chart, sets the legend font color to DarkGray, changes the legend area fill to a solid type with 50 % transparency, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart
                Chart chart = sheet.Charts[0];

                // ----- Legend Font Color -----
                if (chart.Legend != null)
                {
                    // Set legend font color to dark gray
                    chart.Legend.Font.Color = Color.DarkGray;

                    // Set legend fill to solid with 50% transparency
                    chart.Legend.Area.FillFormat.FillType = FillType.Solid;
                    chart.Legend.Area.FillFormat.Transparency = 0.5; // 0 = opaque, 1 = fully transparent
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
