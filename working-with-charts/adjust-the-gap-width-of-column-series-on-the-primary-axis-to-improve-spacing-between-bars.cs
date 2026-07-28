// Title: C# – Adjust Column Series Gap Width in Aspose.Cells Chart to Increase Bar Spacing
// Description: This example creates a new workbook, adds a 2‑D column chart from sample data, and sets the NSeries.GapWidth property to 150 % to widen the space between bars before saving the file as AdjustedGapWidth.xlsx.
// Keywords: Aspose.Cells | C# | column chart | gap width | bar spacing | Chart.NSeries.GapWidth | Excel chart formatting | Aspose.Cells example | GitHub source | chart customization
// Common Searches: Aspose.Cells change column series gap width | increase spacing between bars in Aspose.Cells chart | C# set GapWidth property Aspose.Cells | adjust column chart bar spacing .NET | how to modify chart series gap width Aspose
// Developer Intent: Modify the GapWidth of a column series to control bar spacing in an Aspose.Cells chart.
// Use Cases: Create column charts with clearer visual separation for presentations. | Programmatically adapt bar spacing when the number of categories changes. | Generate Excel reports where column chart readability is critical.
// AI Prompts: Generate C# code that sets GapWidth for each series in a stacked column chart using Aspose.Cells. | Explain how to calculate an optimal GapWidth based on category count and apply it in Aspose.Cells. | Show how to retrieve and modify GapWidth for existing charts in a workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, adds a 2‑D column chart from sample data, and sets the NSeries.GapWidth property to 150 % to widen the space between bars before saving the file as AdjustedGapWidth.xlsx.
    public class AdjustColumnSeriesGapWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the column chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a 2‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Adjust the gap width of the first (and only) series on the primary axis
                // GapWidth is a percentage of the column width (0‑500). Larger values increase spacing.
                chart.NSeries[0].GapWidth = 150; // 150% of the default width

                // Save the workbook
                workbook.Save("AdjustedGapWidth.xlsx");
                Console.WriteLine("Workbook saved as AdjustedGapWidth.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustColumnSeriesGapWidth.Run();
        }
    }
}
