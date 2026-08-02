// Title: Aspose.Cells for .NET (C#) – Convert Second Series to Line in a Combo Chart
// Description: A concise C# example that creates an Excel workbook with a column chart, then changes the second data series to a line type to produce a combo chart. The code demonstrates populating data, setting category labels, renaming series, and saving the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET | combo chart Aspose.Cells | change chart series type | line series Excel chart | column to line chart conversion | Excel chart manipulation C# | Aspose.Cells chart API | chart series type example | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells change series to line | C# combo chart column and line | How to modify chart series type after creation Aspose.Cells | Aspose.Cells example for mixed chart types | Create line series in existing column chart .NET
// Developer Intent: Replace the second column series with a line series to build a combo chart using Aspose.Cells for .NET.
// Use Cases: Show monthly sales as bars and profit margin as a line in a single financial report. | Build a dashboard where product quantities are displayed as columns and growth rate as a line to highlight trends. | Create a multi‑axis chart that compares inventory levels (column) with forecasted demand (line) in one view.
// AI Prompts: Generate C# code that adds a secondary Y‑axis to the line series in the combo chart. | Explain how to set different marker styles for the line series while keeping the column series unchanged. | Provide a script to loop through all series in an Aspose.Cells chart and assign types based on a configuration file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartDemo
{
    // A concise C# example that creates an Excel workbook with a column chart, then changes the second data series to a line type to produce a combo chart. The code demonstrates populating data, setting category labels, renaming series, and saving the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a combo chart (initially Column type)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series – both default to Column type
            chart.NSeries.Add("B2:B4", true); // Series 0
            chart.NSeries.Add("C2:C4", true); // Series 1 (the one we will change)

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // Change the second series (index 1) from Column to Line to create a combo effect
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: give the series a name for clarity
            chart.NSeries[0].Name = "Column Series";
            chart.NSeries[1].Name = "Line Series";

            // Save the workbook
            workbook.Save("ComboChart_Output.xlsx");
        }
    }
}
