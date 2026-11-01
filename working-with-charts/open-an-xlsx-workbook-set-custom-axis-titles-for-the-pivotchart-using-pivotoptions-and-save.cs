// Title: Add custom X‑axis and Y‑axis titles to a PivotChart in an existing XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to change the X axis title of a PivotChart using Aspose.Cells C# | Aspose.Cells set custom category axis label on an existing Excel PivotChart | C# update value axis title of a PivotChart in an XLSX file with Aspose.Cells | modify pivot chart axis titles programmatically with Aspose.Cells for .NET | load workbook and edit PivotChart axis labels using Aspose.Cells API
// Tags: Aspose.Cells pivot chart axis title modification | C# set custom axis titles in Excel PivotChart | Aspose.Cells edit chart axis labels XLSX | programmatic pivot chart axis customization .NET | load workbook modify pivot chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX workbook, accesses the first chart (assumed to be a PivotChart), assigns custom text to the CategoryAxis and ValueAxis titles, and saves the workbook, with basic error handling for missing files or charts.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the PivotChart
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart (expected to be a PivotChart)
            Chart pivotChart = sheet.Charts[0];

            // Set custom titles for the X (Category) and Y (Value) axes
            pivotChart.CategoryAxis.Title.Text = "Custom X Axis Title";
            pivotChart.ValueAxis.Title.Text = "Custom Y Axis Title";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
