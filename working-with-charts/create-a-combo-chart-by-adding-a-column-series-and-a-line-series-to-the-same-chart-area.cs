// Title: Create a Column‑Line Combo Chart with Aspose.Cells for .NET (C#)
// Description: This C# example uses Aspose.Cells for .NET to generate an Excel workbook, fill quarterly sales and profit data, and build a combo chart that combines a column series (sales) and a line series (profit) on a shared category axis, then saves the file as ComboChart.xlsx.
// Keywords: Aspose.Cells | C# | combo chart | column chart | line chart | mixed chart | Excel chart automation | set series type | chart series range | Excel export | Aspose.Cells example
// Common Searches: how to create a combo chart with column and line series using Aspose.Cells C# | Aspose.Cells set different chart types for multiple series | add line series to an existing column chart Aspose.Cells | define category axis data for a combo chart in Aspose.Cells | Aspose.Cells mixed chart example .NET
// Developer Intent: Generate an Excel workbook that contains a combo chart where one series is displayed as columns and another as a line, using Aspose.Cells for .NET.
// Use Cases: Business reports that need sales shown as columns and profit shown as a line on the same chart. | Dashboard widgets where different metrics require distinct visual styles. | Automated financial statements that export data to Excel with a combined column‑line chart for stakeholder presentations.
// AI Prompts: Show how to change the line series color and marker style in the combo chart with Aspose.Cells. | Provide code to add a secondary Y‑axis for the line series in the combo chart. | Explain how to bind chart data ranges from a DataTable and create a combo chart dynamically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example uses Aspose.Cells for .NET to generate an Excel workbook, fill quarterly sales and profit data, and build a combo chart that combines a column series (sales) and a line series (profit) on a shared category axis, then saves the file as ComboChart.xlsx.
class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A – categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Column B – values for the column series
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["B5"].PutValue(210);

        // Column C – values for the line series
        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(55);
        sheet.Cells["C5"].PutValue(70);

        // Add a chart (initially a Column chart) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 22, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add the first series (column series) and set its type explicitly
        chart.NSeries.Add("=Sheet1!$B$2:$B$5", true);
        chart.NSeries[0].Type = ChartType.Column; // column series

        // Add the second series (line series) and set its type to Line
        chart.NSeries.Add("=Sheet1!$C$2:$C$5", true);
        chart.NSeries[1].Type = ChartType.Line; // line series

        // Set the category (X‑axis) data for both series
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

        // Optional: give the series meaningful names
        chart.NSeries[0].Name = "Sales";
        chart.NSeries[1].Name = "Profit";

        // Save the workbook
        workbook.Save("ComboChart.xlsx");
    }
}
