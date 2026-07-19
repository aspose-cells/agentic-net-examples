// Title: Aspose.Cells C# – Show Only Series Names in Chart Legend (Hide Category Entries)
// Description: Creates a workbook, adds sample data for two series across three categories, inserts a column chart, and configures the legend so that only the series names appear. The code sets ShowLegend = true and ensures each series' LegendEntry is visible while no category entries are shown.
// Keywords: Aspose.Cells chart legend | C# hide category legend entries | show series only legend Aspose.Cells | Excel chart legend customization | IsDeleted LegendEntry Aspose.Cells | NSeries legend C# | column chart legend series only | Aspose.Cells API legend settings
// Common Searches: Aspose.Cells show only series names in legend | C# hide category entries from chart legend Aspose.Cells | How to display series only in Excel chart legend using Aspose.Cells | Aspose.Cells legend entry visibility C# | Remove category labels from chart legend Aspose.Cells
// Developer Intent: Configure an Aspose.Cells chart so the legend lists only the series names and excludes any category or point entries.
// Use Cases: Generate Excel reports with column charts where the legend is limited to series identifiers for a cleaner layout. | Build dashboards that already label axes, requiring the legend to show only series names. | Programmatically modify existing charts to suppress category or point legend entries while keeping series entries visible.
// AI Prompts: Write C# code with Aspose.Cells that creates a line chart and configures the legend to display only series names, removing all category entries. | Provide an Aspose.Cells example that iterates over NSeries and sets each series' LegendEntry.IsDeleted property to control legend content. | Explain how to hide point legend entries and keep series legend entries visible in Aspose.Cells charts using the LegendEntry API.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // Creates a workbook, adds sample data for two series across three categories, inserts a column chart, and configures the legend so that only the series names appear. The code sets ShowLegend = true and ensures each series' LegendEntry is visible while no category entries are shown.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: two series (Series1, Series2) across three categories (A, B, C)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series1 values, categories taken from A2:A4
                chart.NSeries.Add("C2:C4", true); // Series2 values, same categories

                // Ensure the legend is displayed
                chart.ShowLegend = true;

                // Keep only series legend entries; point legend entries are not applicable for column charts
                foreach (Series series in chart.NSeries)
                {
                    // Ensure the series legend entry is visible
                    series.LegendEntry.IsDeleted = false;
                }

                // Save the workbook
                workbook.Save("ChartLegendSeriesOnly.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
