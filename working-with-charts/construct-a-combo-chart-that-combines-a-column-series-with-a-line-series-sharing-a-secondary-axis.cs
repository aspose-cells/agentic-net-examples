// Title: Aspose.Cells .NET: Create a Combo Chart with Column and Line Series on a Secondary Axis
// Description: This example shows how to generate an Excel workbook, populate it with month‑wise sales and profit data, add a column chart, convert a second series to a line chart, plot that line on a secondary Y‑axis, set a custom axis title, and save the file as an .xlsx document using Aspose.Cells for C#.
// Keywords: Aspose.Cells | combo chart .NET | column and line series | secondary axis chart | C# Excel chart example | multiple axes Aspose.Cells | chart customization Aspose.Cells | Excel combo chart code | Aspose.Cells ChartType.Line | Aspose.Cells ChartType.Column
// Common Searches: Aspose.Cells create combo chart with secondary axis | C# add line series to column chart Aspose.Cells | how to plot chart series on secondary Y axis using Aspose.Cells | Aspose.Cells set secondary axis title | combo chart example Aspose.Cells .NET
// Developer Intent: Generate an Excel combo chart that combines a column series and a line series, with the line series displayed on a secondary Y‑axis.
// Use Cases: Financial reports that need sales (columns) and profit (line) on different scales. | Dashboard worksheets where two metrics require separate axes for clear comparison. | Automated Excel exports that include multi‑axis visualizations for stakeholder presentations.
// AI Prompts: Write C# code with Aspose.Cells to build a combo chart: column series for sales, line series for profit, plot the line on a secondary axis, add a secondary axis title, and save as .xlsx. | Explain how to change a series type to Line and enable PlotOnSecondAxis in Aspose.Cells. | Provide step‑by‑step instructions to customize the secondary value axis title in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartDemo
{
    // This example shows how to generate an Excel workbook, populate it with month‑wise sales and profit data, add a column chart, convert a second series to a line chart, plot that line on a secondary Y‑axis, set a custom axis title, and save the file as an .xlsx document using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories (X‑axis)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column series values (primary Y axis)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(170);
            sheet.Cells["B5"].PutValue(200);

            // Line series values (secondary Y axis)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);
            sheet.Cells["C5"].PutValue(70);

            // Add a combo chart (base type Column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the column series (first series)
            chart.NSeries.Add("B2:B5", true);               // Values
            chart.NSeries.CategoryData = "A2:A5";           // Categories

            // Add the line series (second series)
            chart.NSeries.Add("C2:C5", true);               // Values

            // Change the second series to a line chart
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the line series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize secondary axis title
            chart.SecondValueAxis.Title.Text = "Profit (Secondary Axis)";

            // Save the workbook
            workbook.Save("ComboChart_ColumnLine_SecondaryAxis.xlsx");
        }
    }
}
