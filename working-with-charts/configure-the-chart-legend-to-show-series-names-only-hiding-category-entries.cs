// Title: Show only series names in a column chart legend while hiding category entries using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that removes category legend entries from a column chart, leaving only the series names visible. | Demonstrate how to iterate over LegendEntryCollection in Aspose.Cells to delete non‑series legend items and keep the legend displayed.
// Common Searches: Aspose.Cells C# hide category items in chart legend | remove category entries from Excel chart legend using Aspose.Cells | display only series names in column chart legend Aspose.Cells .NET | how to delete non‑series legend entries in Aspose.Cells chart | Aspose.Cells legend customization hide categories
// Tags: Aspose.Cells legend entry deletion | C# chart legend series only | Aspose.Cells column chart legend customization | Excel chart legend hide categories Aspose | Aspose.Cells LegendEntryCollection usage

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // The example creates a workbook with sample data, adds a column chart, and then iterates over the chart's LegendEntryCollection to delete entries that correspond to categories, ensuring that only the series names appear in the legend before saving the file as ChartLegendSeriesOnly.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data:
            // Column A – Categories, Column B – Series 1, Column C – Series 2
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

            // Set data ranges – two series sharing the same categories
            chart.NSeries.Add("B2:B4", true);   // Series 1 values, categories taken from A2:A4
            chart.NSeries.Add("C2:C4", true);   // Series 2 values
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the chart is calculated so legend entries are generated
            chart.Calculate();

            // The legend may contain entries for both series and categories (e.g., for certain chart types).
            // To keep only series names, delete legend entries that are not associated with a series.
            // Series legend entries are the first N entries where N = number of series.
            LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
            int seriesCount = chart.NSeries.Count;

            for (int i = seriesCount; i < legendEntries.Count; i++)
            {
                // Hide (delete) category‑related legend entries
                legendEntries[i].IsDeleted = true;
            }

            // Optionally, ensure the legend itself is visible
            chart.ShowLegend = true;

            // Save the workbook
            workbook.Save("ChartLegendSeriesOnly.xlsx");
        }
    }
}
