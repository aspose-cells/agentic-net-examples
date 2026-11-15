// Title: Convert an existing Excel chart to a 3‑D Cone chart with Aspose.Cells for .NET (depth adjustment not available)
// Common Searches: c# aspnet change excel chart type to 3d cone using Aspose.Cells | how to set 3d chart depth in Aspose.Cells .NET | Aspose.Cells example for converting a chart to a cone shape in C# | programmatically modify Excel chart perspective with Aspose.Cells | convert existing workbook chart to 3d cone chart C# Aspose.Cells
// Tags: Aspose.Cells convert chart to 3D cone | C# change Excel chart type Aspose.Cells | Aspose.Cells 3D chart depth limitation | modify chart perspective Aspose.Cells .NET | set chart type Cone Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook, checks for a chart on the first worksheet, changes the chart's type to a 3‑D Cone by assigning ChartType.Cone, notes that the DepthPercent property is not available in the current Aspose.Cells version, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Change the chart type to a 3‑D Cone chart
            chart.Type = ChartType.Cone;

            // Note: DepthPercent property is not available in the current Aspose.Cells version.
            // If needed, additional 3‑D formatting can be applied using other available properties.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
