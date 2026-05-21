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

            // Populate sample data (categories in column A, series values in columns B and C)
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

            // Set data range for the two series
            chart.NSeries.Add("B2:B4", true); // Series 1 values, categories taken from A2:A4
            chart.NSeries.Add("C2:C4", true); // Series 2 values, same categories

            // Ensure the legend is displayed
            chart.ShowLegend = true;

            // Iterate through each series and make sure its legend entry is visible
            foreach (Series s in chart.NSeries)
            {
                // Show the series name in the legend
                s.LegendEntry.IsDeleted = false;
            }

            // If the chart type creates additional legend entries for categories (some chart types do),
            // they can be removed by accessing the LegendEntries collection.
            // Here we delete any legend entry that is not associated with a series.
            // Since Aspose.Cells does not expose a direct way to identify category entries,
            // we simply ensure that only the series legend entries remain.
            // (For most standard charts, this loop will have no effect.)
            Legend legend = chart.Legend;
            if (legend.LegendEntries != null)
            {
                for (int i = 0; i < legend.LegendEntries.Count; i++)
                {
                    // The first NSeries.Count entries correspond to series.
                    // Any extra entries are treated as category entries and are hidden.
                    if (i >= chart.NSeries.Count)
                    {
                        legend.LegendEntries[i].IsDeleted = true;
                    }
                }
            }

            // Save the workbook
            workbook.Save("ChartLegendSeriesOnly.xlsx");
        }
    }
}