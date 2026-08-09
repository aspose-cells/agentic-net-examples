// Title: Aspose.Cells for .NET – Link Chart Data Labels to Cells for Dynamic Number‑Format Inheritance
// Description: Shows how to format cells with a custom pattern, attach a column chart’s data labels to those cells, and automatically inherit the cell’s number format via the NumberFormatLinked property.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | linked source cells | number format inheritance | custom number format | dynamic chart labeling | column chart
// Common Searches: Aspose.Cells link chart data labels to cells | inherit number format from worksheet cells in chart labels | dynamic data label formatting Aspose.Cells C# | NumberFormatLinked property example | chart label custom format using linked source
// Developer Intent: Connect chart data labels to worksheet cells so they automatically adopt the cells' number format.
// Use Cases: Display values with units on column‑chart labels by linking to cells formatted as "#,##0.00 \"units\"". | Update label appearance instantly when the source cell’s format changes, removing manual label edits. | Create reusable chart templates where label styling is driven by cell formatting, enabling consistent reporting across projects.
// AI Prompts: Write C# code with Aspose.Cells that links a chart’s data labels to a cell range and enables NumberFormatLinked. | Explain the role of NumberFormatLinked and how to set up a linked source for custom label formatting. | Adapt the example to a line chart and use a percentage format like "0.0%" for the linked label cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDynamicDataLabelFormatting
{
    // Shows how to format cells with a custom pattern, attach a column chart’s data labels to those cells, and automatically inherit the cell’s number format via the NumberFormatLinked property.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1234.56);
            sheet.Cells["B3"].PutValue(7890.12);

            // Cells that will provide the formatted labels
            sheet.Cells["C1"].PutValue("Formatted Value");
            sheet.Cells["C2"].PutValue(1234.56);
            sheet.Cells["C3"].PutValue(7890.12);

            // Apply a custom number format to the label source cells (e.g., show units)
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00 \"units\"";
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            sheet.Cells.CreateRange("C2:C3").ApplyStyle(style, flag);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set data range for the series
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Configure data labels to use the formatted cells
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.LinkedSource = "C2:C3";          // Link to formatted cells
            series.DataLabels.NumberFormatLinked = true;      // Inherit number format from linked cells

            // Save the workbook
            workbook.Save("DynamicDataLabelFormatting.xlsx");
        }
    }
}
