// Title: Format Z‑Axis (Series Axis) Tick Labels as Percentages in a 3‑D Column Chart – Aspose.Cells for .NET
// Description: The example builds a new workbook, fills it with numeric data, adds a 3‑D column chart, and applies a percentage number format (0%) to the Series (Z) axis tick labels before saving the file.
// Keywords: Aspose.Cells | C# | .NET | 3D column chart | SeriesAxis | Z axis | TickLabels | NumberFormat | percentage format | Excel chart formatting | chart axis customization
// Common Searches: Aspose.Cells set Z axis label to percent | C# format series axis tick labels as percentage | 3D chart axis number format Aspose.Cells | how to display Z axis values as % in Excel using .NET | change chart axis format to percent Aspose
// Developer Intent: Apply a percentage number format to the Z‑axis (Series axis) tick labels of a 3‑D chart with Aspose.Cells.
// Use Cases: Building financial dashboards where the series axis must show relative percentages. | Generating automated Excel reports that compare product shares using a 3‑D column chart. | Creating Excel visualizations for market‑share analysis with percentage labels on the Z‑axis.
// AI Prompts: Generate C# code that creates a 3‑D column chart with Aspose.Cells and sets the SeriesAxis tick label format to "0%". | Show how to format the Z‑axis of a 3‑D chart as percentages in Aspose.Cells, including workbook saving and folder handling. | Explain the steps to change the number format of a chart's Series axis in Aspose.Cells and list supported format strings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example builds a new workbook, fills it with numeric data, adds a 3‑D column chart, and applies a percentage number format (0%) to the Series (Z) axis tick labels before saving the file.
class SetZAxisLabelAsPercentage
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a 3‑D chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B2"].PutValue(0.1); // 10%
            sheet.Cells["B3"].PutValue(0.2); // 20%
            sheet.Cells["B4"].PutValue(0.3); // 30%

            sheet.Cells["C2"].PutValue(0.4); // 40%
            sheet.Cells["C3"].PutValue(0.5); // 50%
            sheet.Cells["C4"].PutValue(0.6); // 60%

            // Add a 3‑D column chart (has a Z axis)
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // For 3‑D charts the Z axis corresponds to the Series axis
            // Set its tick label format to display values as percentages
            chart.SeriesAxis.TickLabels.NumberFormat = "0%";

            // Define output file path
            string outputPath = "ZAxisPercentage.xlsx";

            // Ensure the output directory exists (handle possible null)
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
