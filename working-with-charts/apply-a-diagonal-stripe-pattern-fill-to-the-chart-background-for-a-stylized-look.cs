// Title: How to apply a diagonal stripe pattern fill to an Excel chart background using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook in C# with Aspose.Cells, add a column chart, and set the chart area fill to a LightDownwardDiagonal pattern with custom foreground and background colors. | Programmatically style a chart's background in Aspose.Cells by applying a diagonal stripe pattern fill and save the workbook as an .xlsx file. | Generate a column chart using Aspose.Cells and configure its ChartArea.FillFormat to use a LightDownwardDiagonal pattern with specified colors.
// Common Searches: Aspose.Cells C# set chart area pattern fill diagonal stripes | apply LightDownwardDiagonal fill to Excel chart background using Aspose.Cells | C# example for chart background pattern styling with Aspose.Cells | how to change chart area fill type to pattern in Aspose.Cells .NET | set foreground and background colors for chart pattern fill Aspose.Cells
// Tags: Aspose.Cells chart area pattern fill | C# diagonal stripe fill Aspose.Cells | ChartArea.FillFormat pattern FillType Aspose | Excel chart background styling .NET | LightDownwardDiagonal fill Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a new workbook, adds sample data, inserts a column chart, and applies a LightDownwardDiagonal pattern fill with light blue foreground and dark blue background to the chart area, then saves the file as DiagonalStripeChart.xlsx.
class DiagonalStripeChartBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a diagonal stripe pattern to the chart area background
        chart.ChartArea.Area.FillFormat.FillType = FillType.Pattern;
        chart.ChartArea.Area.FillFormat.PatternFill.Pattern = FillPattern.LightDownwardDiagonal;
        chart.ChartArea.Area.FillFormat.PatternFill.ForegroundColor = Color.LightBlue;
        chart.ChartArea.Area.FillFormat.PatternFill.BackgroundColor = Color.DarkBlue;

        // Save the workbook
        workbook.Save("DiagonalStripeChart.xlsx", SaveFormat.Xlsx);
    }
}
