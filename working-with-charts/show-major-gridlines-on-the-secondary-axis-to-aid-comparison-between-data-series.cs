// Title: C# – Show Major Gridlines on the Secondary Axis of a Column Chart with Aspose.Cells
// Description: Creates an Excel workbook, adds a column chart with two data series, plots the second series on the secondary value axis, and makes the secondary axis major gridlines visible (optionally blue). The workbook is saved as an .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# chart secondary axis | Excel column chart gridlines | major gridlines secondary axis | secondary value axis styling | chart gridline color | plot series on second axis | Excel automation C# | Aspose.Cells example
// Common Searches: Aspose.Cells show secondary axis gridlines C# | enable major gridlines on secondary value axis Aspose.Cells | column chart secondary axis styling Aspose.Cells .NET | how to make secondary axis gridlines visible in Excel using C# | Aspose.Cells sample for secondary axis gridlines
// Developer Intent: Display and style major gridlines on a chart's secondary value axis.
// Use Cases: Align two data series with different scales by adding visible gridlines to the secondary axis. | Generate Excel reports where secondary‑axis gridlines improve readability of comparative charts. | Automate dashboard creation that requires custom styling (color, thickness) of secondary axis gridlines.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line chart, plots a series on the secondary axis, and sets the secondary axis major gridlines to visible with a custom color. | Show how to toggle secondary axis major gridlines on or off based on a boolean variable in an Aspose.Cells workbook. | Explain how to change the line style, thickness, and spacing of major gridlines on the secondary value axis using Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates an Excel workbook, adds a column chart with two data series, plots the second series on the secondary value axis, and makes the secondary axis major gridlines visible (optionally blue). The workbook is saved as an .xlsx file using Aspose.Cells for .NET.
class ShowSecondaryAxisMajorGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["C3"].PutValue(200);
        sheet.Cells["C4"].PutValue(300);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series and set category data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Show major gridlines on the secondary value axis
        chart.SecondValueAxis.MajorGridLines.IsVisible = true;
        chart.SecondValueAxis.MajorGridLines.Color = Color.Blue; // optional styling

        // Save the workbook
        workbook.Save("SecondaryAxisMajorGridlines.xlsx");
    }
}
