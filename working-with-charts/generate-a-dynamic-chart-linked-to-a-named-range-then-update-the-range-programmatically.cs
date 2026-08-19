// Title: Create a Dynamic Column Chart Linked to a Named Range and Update It Programmatically – Aspose.Cells for .NET (C#)
// Description: This example shows how to generate a new workbook, populate it with categories and values, define a named range (DataRange), attach a column chart to that range using SetChartDataRange, add new rows at runtime, expand the named range, recalculate the chart, and save the file as an XLSX document.
// Keywords: Aspose.Cells | dynamic chart | named range | C# | .NET | SetChartDataRange | chart refresh | extend range programmatically | Excel automation | column chart
// Common Searches: Aspose.Cells bind chart to named range C# | update chart data range programmatically Aspose.Cells | extend named range for Excel chart .NET | refresh Aspose.Cells chart after adding rows | SetChartDataRange example with named range
// Developer Intent: Generate a column chart that references a named range and automatically reflects data added later through code.
// Use Cases: Monthly sales dashboard where new month values automatically appear in the chart. | Financial model template that expands chart ranges as users append forecast rows. | Reporting tool that updates visualizations without recreating charts after data imports.
// AI Prompts: Show C# code to link an Aspose.Cells column chart to a named range and refresh it after extending the range. | How can I programmatically add rows, update a named range, and recalculate a chart using Aspose.Cells for .NET? | Explain the correct syntax for SetChartDataRange when the source is a named range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartDemo
{
    // This example shows how to generate a new workbook, populate it with categories and values, define a named range (DataRange), attach a column chart to that range using SetChartDataRange, add new rows at runtime, expand the named range, recalculate the chart, and save the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10); // 20,30,40,50
            }

            // ---------- Define a named range that will be used by the chart ----------
            int nameIndex = workbook.Worksheets.Names.Add("DataRange");
            Name dataRangeName = workbook.Worksheets.Names[nameIndex];
            // RefersTo must start with '=' and include sheet name
            dataRangeName.RefersTo = $"={sheet.Name}!$B$2:$B$5";

            // ---------- Add a column chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // ---------- Link the chart to the named range ----------
            // SetChartDataRange expects a range address without the leading '='
            string area = dataRangeName.RefersTo.Substring(1); // "Sheet1!$B$2:$B$5"
            chart.SetChartDataRange(area, true); // true => plot by column

            // Optional: set category axis (using column A)
            chart.NSeries.CategoryData = "A2:A5";

            // ---------- Update the underlying range programmatically ----------
            // Add a new data point
            sheet.Cells["A6"].PutValue("Item 5");
            sheet.Cells["B6"].PutValue(60);

            // Extend the named range to include the new row
            dataRangeName.RefersTo = $"={sheet.Name}!$B$2:$B$6";

            // Refresh the chart so it picks up the extended range
            chart.Calculate();

            // ---------- Save the workbook ----------
            workbook.Save("DynamicChartLinkedToNamedRange.xlsx", SaveFormat.Xlsx);
        }
    }
}
