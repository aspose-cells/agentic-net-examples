// Title: Remove legend text fill from all charts in an Aspose.Cells workbook (C#)
// Description: Creates a workbook with multiple charts, loops through every worksheet and chart, accesses each LegendEntry, sets IsTextNoFill to true, and saves the file. Demonstrates a global legend‑style change using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# legend no fill | chart legend text transparent Aspose | set IsTextNoFill all charts | loop through worksheets charts Aspose.Cells | remove legend background fill .NET
// Common Searches: how to clear legend text fill for all charts Aspose.Cells | C# loop through workbook charts set legend no fill | Aspose.Cells remove legend background from multiple charts | global legend formatting Aspose.Cells C#
// Developer Intent: Apply a transparent background to legend text on every chart in a workbook.
// Use Cases: Standardize legend appearance in financial dashboards with dozens of charts. | Prepare a workbook for PDF export where legend fill interferes with readability. | Retrofit existing reports to ensure consistent, fill‑free legend text across all sheets.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all worksheets and charts to set LegendEntry.IsTextNoFill = true. | Explain how to safely handle charts that may not have legend entries when clearing legend fill in Aspose.Cells. | Describe the visual impact of the IsTextNoFill property on chart legends and how to apply it globally.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with multiple charts, loops through every worksheet and chart, accesses each LegendEntry, sets IsTextNoFill to true, and saves the file. Demonstrates a global legend‑style change using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(100);
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(60);
        sheet.Cells["C3"].PutValue(120);

        // Add first chart (Column)
        int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart1 = sheet.Charts[chartIdx1];
        chart1.NSeries.Add("B2:C3", true);
        chart1.NSeries.CategoryData = "A2:A3";

        // Add second chart (Pie)
        int chartIdx2 = sheet.Charts.Add(ChartType.Pie, 16, 0, 26, 5);
        Chart chart2 = sheet.Charts[chartIdx2];
        chart2.NSeries.Add("B2:B3", true);
        chart2.NSeries.CategoryData = "A2:A3";

        // Loop through all worksheets
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Loop through all charts in the worksheet
            foreach (Chart ch in ws.Charts)
            {
                // Get the collection of legend entries (may be null for some chart types)
                LegendEntryCollection entries = ch.Legend.LegendEntries;
                if (entries != null)
                {
                    // Loop through each legend entry and set no fill for the text
                    foreach (LegendEntry entry in entries)
                    {
                        entry.IsTextNoFill = true;
                    }
                }
            }
        }

        // Save the workbook
        workbook.Save("AllChartsLegendNoFill.xlsx", SaveFormat.Xlsx);
    }
}
