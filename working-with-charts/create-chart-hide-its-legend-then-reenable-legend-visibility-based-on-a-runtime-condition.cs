// Title: Create a column chart in Aspose.Cells for .NET, hide the legend, and reveal it conditionally at runtime
// Common Searches: Aspose.Cells hide chart legend C# example | show chart legend conditionally with Aspose.Cells .NET | toggle ShowLegend property based on runtime value in Aspose.Cells | C# create column chart and control legend visibility using Aspose.Cells | Aspose.Cells chart legend visibility based on boolean flag
// Tags: Aspose.Cells chart legend hide | conditional legend visibility C# | ShowLegend property Aspose.Cells | column chart creation Aspose.Cells | runtime chart formatting Aspose.Cells | Excel workbook chart legend toggle

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsChartExample
{
    // The sample creates a new workbook, fills cells A1:A5 and B1:B5 with sample data, adds a column chart, initially hides its legend using the ShowLegend property, evaluates a runtime boolean condition, and re‑enables the legend when the condition is true before saving the file as ChartWithConditionalLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // A1:A5 – category axis, B1:B5 – values
                sheet.Cells["A1"].PutValue("Jan");
                sheet.Cells["A2"].PutValue("Feb");
                sheet.Cells["A3"].PutValue("Mar");
                sheet.Cells["A4"].PutValue("Apr");
                sheet.Cells["A5"].PutValue("May");

                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(25);
                sheet.Cells["B5"].PutValue(15);

                // Add a column chart (type, upper‑left row/col, lower‑right row/col)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the series and categories
                chart.NSeries.Add("B1:B5", true);
                chart.NSeries.CategoryData = "A1:A5";

                // Hide legend initially using the correct property
                chart.ShowLegend = false;

                // Determine at runtime whether to show the legend
                bool showLegend = GetShowLegendCondition();

                // Show legend if condition is met
                if (showLegend)
                {
                    chart.ShowLegend = true;
                }

                // Save the workbook
                string outputPath = "ChartWithConditionalLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Example runtime condition – replace with real logic as needed
        static bool GetShowLegendCondition()
        {
            // Toggle based on current second (even seconds show legend)
            return DateTime.Now.Second % 2 == 0;
        }
    }
}
