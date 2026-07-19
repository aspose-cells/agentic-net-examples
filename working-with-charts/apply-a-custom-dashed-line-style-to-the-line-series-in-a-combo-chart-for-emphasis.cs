// Title: Aspose.Cells .NET – Create a Combo Chart with a Red Dashed Line Series (C#)
// Description: This C# example builds a new workbook, fills it with monthly sales and target data, adds a combo chart (column for sales, line for target), converts the second series to a line type, and formats that line with a red medium‑weight dashed border and circular markers. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | combo chart | line series | dashed line style | chart border formatting | Excel chart programmatically | MsoLineDashStyle.Dash | WeightType.MediumLine | ChartMarkerType.Circle
// Common Searches: Aspose.Cells set dashed line for chart series | C# combo chart column and line Aspose.Cells | how to change line series border color in Aspose.Cells | customize chart markers Aspose.Cells .NET | apply dash style to Excel chart using Aspose
// Developer Intent: Add a line series to a combo chart and emphasize it with a red dashed border and optional markers.
// Use Cases: Sales vs. target report where the target trend line stands out with a red dashed style. | Financial dashboard that combines column and line series and highlights key trends. | Project timeline where milestone lines are styled with dashed borders for visual distinction.
// AI Prompts: Generate C# code with Aspose.Cells to create a column‑line combo chart and format the line series as a red medium‑weight dashed line. | Show how to set marker style and border dash type for a line series in an Aspose.Cells chart. | Explain the steps to change dash type, color, and weight of a chart series border using Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# example builds a new workbook, fills it with monthly sales and target data, adds a combo chart (column for sales, line for target), converts the second series to a line type, and formats that line with a red medium‑weight dashed border and circular markers. The workbook is saved as an Excel file.
class ComboChartWithDashedLineSeries
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(210);

            sheet.Cells["C1"].PutValue("Target");
            sheet.Cells["C2"].PutValue(130);
            sheet.Cells["C3"].PutValue(140);
            sheet.Cells["C4"].PutValue(190);
            sheet.Cells["C5"].PutValue(200);

            // Add a combo chart (initially a column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the column series (first series)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Sales";

            // Add the line series (second series) and set its type to Line
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Target";
            chart.NSeries[1].Type = ChartType.Line; // Convert to line series

            // Set the category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Apply a custom dashed line style to the line series for emphasis
            chart.NSeries[1].Border.DashType = MsoLineDashStyle.Dash; // Dashed line
            chart.NSeries[1].Border.Color = Color.Red;                // Red color
            chart.NSeries[1].Border.Weight = WeightType.MediumLine;   // Medium thickness

            // Optional: make the line series markers more visible
            chart.NSeries[1].Marker.MarkerStyle = ChartMarkerType.Circle;
            // Size and Color properties may not be available in some versions; they are optional
            // chart.NSeries[1].Marker.Size = 8;
            // chart.NSeries[1].Marker.Color = Color.Red;

            // Determine output path and ensure directory exists
            string outputFile = Path.Combine(Environment.CurrentDirectory, "ComboChartWithDashedLineSeries.xlsx");
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the combo chart:");
            Console.WriteLine(ex.Message);
        }
    }
}
