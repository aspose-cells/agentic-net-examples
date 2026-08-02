// Title: Aspose.Cells C# – Set Dashed Line Style for Line Series in a Column‑Line Combo Chart
// Description: This C# sample builds a new workbook, inserts sales and profit data, adds a column‑line combo chart, converts the second series to a line type, and highlights it by applying a dash‑dot border, dark‑red color, and medium line weight before saving the workbook as an Excel file.
// Keywords: Aspose.Cells | C# chart formatting | combo chart | line series dash style | MsoLineDashStyle.DashDot | chart series border color | Excel chart styling .NET | custom line weight | Aspose.Cells example | column‑line chart
// Common Searches: how to apply dash style to line series in Aspose.Cells | Aspose.Cells C# set chart series border dash type | combo chart line formatting Aspose.Cells .NET | dash‑dot line series in Excel chart using Aspose | change line series color and weight in Aspose.Cells chart
// Developer Intent: Create a column‑line combo chart and emphasize the line series with a dash‑dot border.
// Use Cases: Highlight profit trends in a sales‑vs‑profit report by using a dark‑red dash‑dot line. | Distinguish multiple line series in a financial dashboard with different dash patterns and weights. | Produce printable Excel charts where the line series stands out for better visual analysis.
// AI Prompts: Write C# code with Aspose.Cells that builds a column‑line combo chart and sets the line series border to DashDot, dark red, medium weight. | Show how to modify an existing Aspose.Cells chart to change the line series dash type, color, and thickness programmatically. | Explain steps to apply various dash styles to several line series in a combo chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsComboChartDemo
{
    // This C# sample builds a new workbook, inserts sales and profit data, adds a column‑line combo chart, converts the second series to a line type, and highlights it by applying a dash‑dot border, dark‑red color, and medium line weight before saving the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column series data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(210);

            // Line series data
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);
            sheet.Cells["C5"].PutValue(80);

            // Add a combo chart (column + line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the column series (first series)
            chart.NSeries.Add("B2:B5", true);
            // Add the line series (second series)
            chart.NSeries.Add("C2:C5", true);
            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Convert the second series to a line type
            chart.NSeries[1].Type = ChartType.Line;

            // Apply custom dashed line style to the line series for emphasis
            // Use DashType property of the series border
            chart.NSeries[1].Border.DashType = MsoLineDashStyle.DashDot;
            chart.NSeries[1].Border.Color = Color.DarkRed;
            chart.NSeries[1].Border.Weight = WeightType.MediumLine;

            // Save the workbook (lifecycle save)
            workbook.Save("ComboChartWithDashedLineSeries.xlsx");
        }
    }
}
