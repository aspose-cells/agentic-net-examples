using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetLegendEntryNoFill
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for demonstration purposes
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(100);

        // Add a chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Loop through all worksheets in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Loop through all charts in the current worksheet
            foreach (Chart ch in ws.Charts)
            {
                // Get the collection of legend entries for the chart
                LegendEntryCollection legendEntries = ch.Legend.LegendEntries;

                // Ensure the collection is not null (e.g., surface charts return null)
                if (legendEntries != null)
                {
                    // Set IsTextNoFill = true for each legend entry (no fill for the text)
                    foreach (LegendEntry entry in legendEntries)
                    {
                        entry.IsTextNoFill = true;
                    }
                }
            }
        }

        // Save the workbook with the updated legend settings
        workbook.Save("AllChartsLegendNoFill.xlsx");
    }
}