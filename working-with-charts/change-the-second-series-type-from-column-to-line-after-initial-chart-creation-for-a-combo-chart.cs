// Title: Change a series to line type in a combo chart with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart with two data series, then convert the second series to a line chart by setting chart.NSeries[1].Type to ChartType.Line, add a title, and save the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | combo chart | change series type | ChartType.Line | NSeries | column chart | line series | Excel chart customization | Aspose.Cells example
// Common Searches: Aspose.Cells change series to line after chart creation | set second series type line combo chart Aspose.Cells .NET | modify chart series type programmatically Aspose.Cells | create column and line combo chart C# Aspose.Cells | how to use NSeries.Type property Aspose.Cells
// Developer Intent: Update the chart type of a specific series in an existing Aspose.Cells combo chart, converting it from column to line.
// Use Cases: Sales dashboard: display monthly sales as columns and profit margin as a line. | Financial report: show revenue bars with a line for year‑over‑year growth. | Performance tracking: plot daily output as columns and average trend as a line.
// AI Prompts: Provide C# code that changes the second series of an Aspose.Cells combo chart to a line type after the chart is created. | Show an example of building a column‑plus‑line combo chart with Aspose.Cells for .NET. | Explain how to assign different ChartType values to multiple NSeries in a single Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartDemo
{
    // Demonstrates how to create a workbook, add a column chart with two data series, then convert the second series to a line chart by setting chart.NSeries[1].Type to ChartType.Line, add a title, and save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // First series (Column)
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Second series (will be changed to Line)
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a combo chart (initially a Column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.CategoryData = "A2:A4";

            // Change the second series type from Column to Line
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: give the chart a title
            chart.Title.Text = "Combo Chart (Column + Line)";

            // Save the workbook
            workbook.Save("ComboChart_Output.xlsx");
        }
    }
}
