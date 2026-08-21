// Title: Aspose.Cells C# – Show Only Series Names in Chart Legend (Hide Category Entries)
// Description: Creates a workbook, adds a column chart with two series, enables the legend, then uses the LegendEntry collection and the IsDeleted flag to remove any category legend items, leaving only the series names visible before saving the file.
// Keywords: Aspose.Cells chart legend series only | hide category legend entries C# | Aspose.Cells LegendEntry.IsDeleted | customize chart legend Aspose.Cells | C# Aspose.Cells chart legend filtering
// Common Searches: Aspose.Cells show only series names in legend | remove category entries from chart legend C# | Aspose.Cells hide legend categories | how to filter legend entries Aspose.Cells | C# chart legend customization Aspose.Cells
// Developer Intent: Display a chart legend that contains only the series names by programmatically deleting category entries.
// Use Cases: Generate clean reports where the legend should list only data series, not categories. | Automate dashboard creation with consistent legend formatting across multiple chart types. | Programmatically adjust legend content after adding dynamic series to a workbook.
// AI Prompts: Write C# code with Aspose.Cells to hide all category entries in a line chart legend while keeping series entries visible. | Explain the purpose of LegendEntry.IsDeleted and how to calculate the number of series to retain in a chart legend. | Suggest an alternative method to exclude category items from a chart legend without iterating over LegendEntryCollection.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // Creates a workbook, adds a column chart with two series, enables the legend, then uses the LegendEntry collection and the IsDeleted flag to remove any category legend items, leaving only the series names visible before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (categories in column A, two series in columns B and C)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set data ranges: first series (B2:B4) and second series (C2:C4)
            chart.NSeries.Add("B2:B4", true); // true => use first column as category data
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is displayed
            chart.ShowLegend = true;

            // Configure legend to show only series names.
            // The legend entries collection may contain entries for series and, in some cases, for categories.
            // We keep entries that correspond to series (first N entries) and hide the rest.
            int seriesCount = chart.NSeries.Count;
            LegendEntryCollection legendEntries = chart.Legend.LegendEntries;

            for (int i = 0; i < legendEntries.Count; i++)
            {
                // Keep series legend entries visible
                if (i < seriesCount)
                {
                    legendEntries[i].IsDeleted = false;
                }
                else // Hide any additional entries (e.g., category entries)
                {
                    legendEntries[i].IsDeleted = true;
                }
            }

            // Save the workbook
            workbook.Save("ChartLegendSeriesOnly.xlsx");
        }
    }
}
