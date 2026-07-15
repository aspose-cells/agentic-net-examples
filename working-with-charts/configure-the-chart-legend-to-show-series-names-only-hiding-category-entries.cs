using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendConfiguration
{
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

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1 values
            chart.NSeries.Add("C2:C4", true); // Series 2 values
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is displayed
            chart.ShowLegend = true;

            // Show only series names in the legend:
            // 1. Make sure each series' legend entry is visible.
            // 2. Hide any additional legend entries (e.g., category entries) that may exist.
            int seriesCount = chart.NSeries.Count;
            Legend legend = chart.Legend;
            LegendEntryCollection legendEntries = legend.LegendEntries;

            // Guard against null (some chart types may not support legend entries collection)
            if (legendEntries != null)
            {
                // First, ensure series legend entries are not deleted
                for (int i = 0; i < seriesCount && i < legendEntries.Count; i++)
                {
                    legendEntries[i].IsDeleted = false; // show series name
                }

                // Then, hide any remaining legend entries (commonly category entries)
                for (int i = seriesCount; i < legendEntries.Count; i++)
                {
                    legendEntries[i].IsDeleted = true; // hide category entry
                }
            }

            // Save the workbook
            workbook.Save("ChartLegendSeriesOnly.xlsx");
        }
    }
}