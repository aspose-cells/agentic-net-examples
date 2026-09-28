// Title: Set a custom numeric format on a PivotChart value axis in an existing Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, locates the first PivotChart, and assigns "#,##0.00" to the chart's value axis NumberFormat property. | Create a reusable C# method that takes a file path, chart index, and format string, applies the format to the chart's Y‑axis via Aspose.Cells, and saves the workbook. | Show how to use C# dynamic objects to safely set the ValueAxis.NumberFormat of a chart when the Aspose.Cells API version may differ.
// Common Searches: Aspose.Cells C# set custom number format for pivot chart axis in existing workbook | How to change the value axis format of an Excel chart using Aspose.Cells .NET | Apply numeric format to chart Y axis without affecting data with Aspose.Cells | C# modify chart axis properties using reflection in Aspose.Cells
// Tags: set chart value axis number format Aspose.Cells | custom numeric format for PivotChart .NET | load workbook modify chart Aspose.Cells | dynamic axis property handling C# | apply number format to Excel chart using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, verifies a chart exists on the first worksheet, and uses Aspose.Cells with a dynamic object to set the chart's value (Y) axis NumberFormat to "#,##0.00". It then saves the modified workbook to a new file, handling missing files, absent charts, and formatting exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the chart is on the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            // Get the first chart (replace with appropriate index if necessary)
            Chart chart = sheet.Charts[0];

            // Apply a custom number format to the value (Y) axis labels.
            // Use dynamic to avoid compile‑time binding issues with older API versions.
            try
            {
                dynamic valueAxis = chart.ValueAxis;
                valueAxis.NumberFormat = "#,##0.00";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying number format: {ex.Message}");
            }

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
