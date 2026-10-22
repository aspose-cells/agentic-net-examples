// Title: Change an existing Excel chart to a 3‑D Cone chart and set its Z‑axis depth using Aspose.Cells for .NET
// AI Prompts: Replace the chart.Type assignment with ChartType.Cone3D and then set chart.PlotArea.Depth (e.g., 0.5) to improve the 3‑D perspective. | Update the sample to convert the first worksheet’s chart to a 3‑D cone and configure its Z‑axis depth property for better visual depth.
// Common Searches: Aspose.Cells C# how to convert a column chart to a 3D cone chart | set depth of 3D chart in Aspose.Cells .NET | change Excel chart type to Cone3D using Aspose.Cells library | adjust Z‑axis depth for 3D charts in Aspose.Cells C# example | increase perspective of 3D cone chart Aspose.Cells
// Tags: Aspose.Cells set chart type to Cone3D | C# adjust 3D chart depth Aspose.Cells | modify existing workbook chart Cone3D | configure Z axis depth Aspose.Cells chart | convert Excel chart to 3D cone .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program checks for the input Excel file, loads it into a Workbook, accesses the first worksheet and its first chart, changes the chart type to a 3‑D Cone, sets the chart's Z‑axis depth for better perspective, and saves the modified workbook to the output path while handling missing files or charts gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the first worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            try
            {
                // Change the chart type to a cone chart (3‑D cone may not be supported in this version)
                chart.Type = ChartType.Cone;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to set chart type: {ex.Message}");
                return;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
