// Title: Create a Column‑Line Combo Chart with Aspose.Cells for .NET (C#)
// Description: This example shows how to build an Excel workbook, populate month, sales, and target data, add a column chart, assign category data, then add a second series and switch its type to Line to produce a mixed column‑line combo chart. The workbook is saved as ComboChart.xlsx.
// Keywords: Aspose.Cells | C# | .NET | combo chart | column chart | line chart | mixed chart | Excel chart series | ChartType.Column | ChartType.Line | NSeries | category data
// Common Searches: Aspose.Cells create combo chart C# | add line series to column chart Aspose.Cells | mixed column and line chart example .NET | how to set series type Aspose.Cells | Excel combo chart code sample Aspose
// Developer Intent: Generate an Excel file that contains a combo chart combining a column series and a line series using Aspose.Cells for .NET.
// Use Cases: Compare monthly sales (bars) against a target line in a sales dashboard. | Display revenue bars with profit‑margin trend line for financial reporting. | Show project progress with completed tasks as columns and remaining tasks as a line.
// AI Prompts: Write C# code with Aspose.Cells to create a combo chart that mixes column and line series. | Explain how to assign category data and rename series when building a mixed chart in Aspose.Cells. | Show how to export a workbook containing a combo chart to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartDemo
{
    // This example shows how to build an Excel workbook, populate month, sales, and target data, add a column chart, assign category data, then add a second series and switch its type to Line to produce a mixed column‑line combo chart. The workbook is saved as ComboChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column series values
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(210);

            // Line series values
            sheet.Cells["C1"].PutValue("Target");
            sheet.Cells["C2"].PutValue(130);
            sheet.Cells["C3"].PutValue(140);
            sheet.Cells["C4"].PutValue(170);
            sheet.Cells["C5"].PutValue(200);

            // Add a chart (initially a Column chart) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) data for the chart
            chart.NSeries.CategoryData = "A2:A5";

            // Add the first series (column) using the column values range
            chart.NSeries.Add("B2:B5", true);
            // Force the first series to be displayed as a column
            chart.NSeries[0].Type = ChartType.Column;

            // Add the second series (line) using the line values range
            chart.NSeries.Add("C2:C5", true);
            // Change the second series type to Line to create a combo chart
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: give each series a name (will appear in the legend)
            chart.NSeries[0].Name = "Sales";
            chart.NSeries[1].Name = "Target";

            // Save the workbook with the combo chart
            workbook.Save("ComboChart.xlsx");
        }
    }
}
