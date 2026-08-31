// Title: Set legend entry text to no fill for all charts in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over every worksheet and chart, setting each LegendEntry.IsTextNoFill property to true. | Generate a C# example that creates a workbook with multiple charts and disables the fill of legend text for all charts via a nested foreach loop. | Provide a C# snippet that loops through a workbook's Chart collection and applies LegendEntry.IsTextNoFill = true to remove legend background.
// Common Searches: Aspose.Cells C# remove legend fill from all charts in a workbook | C# loop through charts and set LegendEntry IsTextNoFill property Aspose.Cells | how to disable legend text fill for multiple charts using Aspose.Cells API | iterate over worksheet charts to clear legend background in .NET | set legend entry no fill for column and line charts Aspose.Cells example
// Tags: Aspose.Cells legend entry no fill C# | iterate workbook charts Aspose.Cells | set IsTextNoFill property Aspose.Cells | disable legend text fill .NET Excel | chart legend customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendEntryNoFillDemo
{
    // The example creates a workbook, adds sample data, inserts a column and a line chart, then uses nested foreach loops to traverse every worksheet and each chart. For each chart it accesses the LegendEntries collection and sets the IsTextNoFill property to true, removing the fill from legend text across all charts before saving the file as an XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for charts
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add first chart
            int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart1 = sheet.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B4", true);
            chart1.NSeries.Add("C2:C4", true);
            chart1.NSeries.CategoryData = "A2:A4";

            // Add second chart
            int chartIdx2 = sheet.Charts.Add(ChartType.Line, 16, 0, 26, 5);
            Chart chart2 = sheet.Charts[chartIdx2];
            chart2.NSeries.Add("B2:B4", true);
            chart2.NSeries.Add("C2:C4", true);
            chart2.NSeries.CategoryData = "A2:A4";

            // Loop through all worksheets and their charts
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    // Get the collection of legend entries for the current chart
                    LegendEntryCollection legendEntries = ch.Legend.LegendEntries;

                    // Some chart types (e.g., surface) may return null
                    if (legendEntries != null)
                    {
                        // Set IsTextNoFill = true for each legend entry (no fill for the text)
                        for (int i = 0; i < legendEntries.Count; i++)
                        {
                            legendEntries[i].IsTextNoFill = true;
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save("LegendEntryNoFillDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
