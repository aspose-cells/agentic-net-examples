// Title: Set chart legend background to fully transparent while keeping legend entries fill none using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart legend background transparency to 100 percent | How to hide legend border and make legend background clear in an Excel chart with Aspose.Cells | C# code to make chart legend completely transparent while preserving entry fill using Aspose.Cells | Make Excel chart legend area fill transparent Aspose.Cells .NET example | Remove legend border and set legend area opacity in Aspose.Cells chart
// Tags: chart legend background transparency Aspose.Cells | legend area fill format transparency .NET | hide chart legend border Aspose.Cells | transparent legend background Excel C# | set legend entry fill none Aspose.Cells chart

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses the first chart, sets the legend area fill transparency to 100% for a fully clear background, disables the legend border, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the worksheet
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                    return;
                }

                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Make the legend background transparent by setting full transparency
                chart.Legend.Area.FillFormat.Transparency = 100;

                // Optionally hide the legend border
                chart.Legend.Border.IsVisible = false;

                // Save the modified workbook
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
