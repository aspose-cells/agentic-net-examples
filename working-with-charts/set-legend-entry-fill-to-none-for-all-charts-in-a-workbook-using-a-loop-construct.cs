// Title: Remove legend text fill from all charts in an Aspose.Cells workbook (C#)
// Description: C# example that creates a workbook, adds a column chart, then iterates through every worksheet and each chart to obtain the LegendEntryCollection and set LegendEntry.IsTextNoFill = true, resulting in legend text with no background fill.
// Keywords: Aspose.Cells C# legend no fill | LegendEntry.IsTextNoFill | loop through charts Aspose.Cells | remove legend background color | chart legend formatting C# | Aspose.Cells chart API | Excel legend transparency | batch process chart legends
// Common Searches: Aspose.Cells set legend entry no fill | C# loop through all charts to clear legend fill | how to make legend text transparent in Aspose.Cells | remove legend background color from multiple charts | set IsTextNoFill for chart legends in C#
// Developer Intent: Apply a no‑fill (transparent) style to the text of every legend entry in all charts of a workbook using Aspose.Cells.
// Use Cases: Generate a multi‑chart report where legend text must be invisible for high‑contrast printing. | Batch‑process existing Excel files to clear legend background before converting them to PDF or images. | Create a workbook template that automatically removes legend text fill whenever new charts are added.
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets and charts to set LegendEntry.IsTextNoFill to true. | Show how to safely check for null LegendEntryCollection before applying no‑fill to legend text in each chart. | Explain how to extend the loop to handle charts added dynamically after the initial processing.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendEntryNoFillDemo
{
    // C# example that creates a workbook, adds a column chart, then iterates through every worksheet and each chart to obtain the LegendEntryCollection and set LegendEntry.IsTextNoFill = true, resulting in legend text with no background fill.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(80);

            // Add a chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Loop through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Loop through all charts in the current worksheet
                foreach (Chart ch in ws.Charts)
                {
                    // Get the collection of legend entries; may be null for some chart types
                    LegendEntryCollection legendEntries = ch.Legend.LegendEntries;
                    if (legendEntries != null)
                    {
                        // Iterate over each legend entry and set no fill for the text
                        foreach (LegendEntry entry in legendEntries)
                        {
                            entry.IsTextNoFill = true; // Text will have no fill
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save("AllChartsLegendEntryNoFill.xlsx");
        }
    }
}
