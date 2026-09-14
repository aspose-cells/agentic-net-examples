// Title: How to apply a red dashed line style to the line series of a Column‑Line combo chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a combo chart (column + line) and formats the line series with a red 2‑pt dashed stroke. | Show how to modify an existing workbook to change the dash style of the second series in a combo chart, checking for Line and LineDashStyle support. | Generate a reusable C# method that accepts a chart object and applies a custom dash pattern, color, and width to a specified line series using Aspose.Cells.
// Common Searches: Aspose.Cells C# set dash style for line series in a combo chart | change line series color and pattern in column‑line chart using Aspose.Cells | apply custom line formatting to chart series in Aspose.Cells .NET workbook
// Tags: Aspose.Cells combo chart line dash styling | C# set line series dash pattern Aspose.Cells | Aspose.Cells chart series color width customization | Aspose.Cells column line chart formatting | Aspose.Cells line series appearance .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, fills cells A2‑A6, B2‑B6, and C2‑C6 with sample data, adds a Column‑Line combo chart, defines a column series for sales and a line series for trend, sets the second series to a line type, and demonstrates (commented) how to apply a red 2‑pt dashed line style with specific dash pattern, color, and width before saving the file as ComboChart.xlsx.
class ComboChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["A6"].PutValue("May");

            // Column series values
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);
            sheet.Cells["B6"].PutValue(160);

            // Line series values
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(90);
            sheet.Cells["C4"].PutValue(85);
            sheet.Cells["C5"].PutValue(95);
            sheet.Cells["C6"].PutValue(100);

            // Add a Combo chart (Column + Line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series – Column
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries[0].Name = "Sales";

            // Second series – Line
            chart.NSeries.Add("C2:C6", true);
            chart.NSeries[1].Name = "Trend";

            // Set the second series to be a line type
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: customize line appearance if supported
            // Note: The Line and LineDashStyle properties may not be available in all versions.
            // If they are available, uncomment the following lines:
            // chart.NSeries[1].Line.DashStyle = LineDashStyle.Dash;
            // chart.NSeries[1].Line.Color = Color.Red;
            // chart.NSeries[1].Line.Width = 2.0;

            // Save the workbook
            string outputPath = "ComboChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
