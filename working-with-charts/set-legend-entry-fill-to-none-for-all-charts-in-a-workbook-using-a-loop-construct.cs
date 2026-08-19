// Title: Make all chart legend entries transparent (no fill) in an Aspose.Cells workbook – C# loop
// Description: This C# example creates a workbook, adds a column chart, then iterates through every worksheet and each chart it contains. For each chart it accesses the Legend.LegendEntries collection (skipping charts without legends) and sets the IsTextNoFill property to true, producing legend text without background fill before saving the file as AllChartsLegendNoFill.xlsx.
// Keywords: Aspose.Cells chart legend no fill | C# set legend entry transparent | loop through worksheets charts Aspose | LegendEntry.IsTextNoFill | remove legend background Aspose.Cells | chart formatting programmatically | bulk chart legend styling
// Common Searches: aspocells set legend entry no fill | c# loop all charts remove legend background | how to make chart legend transparent in Aspose.Cells | bulk modify chart legends Aspose.Cells | legend entry IsTextNoFill example
// Developer Intent: Apply a no‑fill style to every legend entry of all charts in a workbook programmatically.
// Use Cases: Standardize legend appearance across dozens of charts in automated financial reports. | Prepare workbooks for high‑resolution PDF export where legend backgrounds must be invisible. | Create a reusable utility that cleans up chart legends before distributing spreadsheets to clients.
// AI Prompts: Generate C# code using Aspose.Cells that loops through all worksheets and charts in a workbook and sets Legend.LegendEntries[i].IsTextNoFill = true for each entry. | Show an example that safely checks for null LegendEntries (e.g., surface charts) and removes the background fill from legends of all other chart types. | Write a method that accepts a Workbook object and applies a transparent legend style to every chart using a foreach loop.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, adds a column chart, then iterates through every worksheet and each chart it contains. For each chart it accesses the Legend.LegendEntries collection (skipping charts without legends) and sets the IsTextNoFill property to true, producing legend text without background fill before saving the file as AllChartsLegendNoFill.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data and a chart (optional, for demonstration)
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("Q1");
        ws.Cells["A3"].PutValue("Q2");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["B2"].PutValue(50);
        ws.Cells["B3"].PutValue(100);

        // Add a chart to the worksheet
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through all charts in the current worksheet
            foreach (Chart ch in sheet.Charts)
            {
                // Get the collection of legend entries; may be null for surface charts
                LegendEntryCollection entries = ch.Legend.LegendEntries;
                if (entries == null) continue;

                // Set IsTextNoFill = true for each legend entry (no fill for the text)
                for (int i = 0; i < entries.Count; i++)
                {
                    entries[i].IsTextNoFill = true;
                }
            }
        }

        // Save the workbook
        workbook.Save("AllChartsLegendNoFill.xlsx");
    }
}
