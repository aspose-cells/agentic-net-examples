// Title: C# – Apply Diagonal Stripe Pattern Fill to Chart Background with Aspose.Cells
// Description: C# example that creates an Excel workbook, adds sample data, inserts a column chart, and sets the chart area to a DarkDownwardDiagonal pattern (light‑gray on white) using Aspose.Cells for .NET, then saves the file as DiagonalStripeChart.xlsx.
// Keywords: Aspose.Cells chart background pattern | C# diagonal stripe fill | FillPattern.DarkDownwardDiagonal | Excel chart styling Aspose | chart area fill type Aspose.Cells | pattern fill example C# | Aspose.Cells GitHub snippet | Excel chart pattern fill
// Common Searches: how to set diagonal stripe pattern for chart background Aspose.Cells C# | Aspose.Cells FillPattern.DarkDownwardDiagonal example | C# code to apply pattern fill to Excel chart area | Aspose.Cells chart area fill type pattern | sample code for chart background pattern fill Aspose
// Developer Intent: Add a diagonal stripe pattern fill to the background of an Excel chart using Aspose.Cells for .NET.
// Use Cases: Design business reports with stylized column charts that feature a light‑gray diagonal stripe background. | Generate Excel workbooks where each chart follows corporate branding by using distinct pattern fills. | Differentiate multiple charts on the same worksheet by applying varied pattern fills to their backgrounds.
// AI Prompts: Show how to change the chart background to a solid color instead of a pattern with Aspose.Cells for .NET. | Provide C# code to apply a horizontal stripe pattern to a chart's plot area using Aspose.Cells. | Explain how to assign different pattern fills to several charts in one worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// C# example that creates an Excel workbook, adds sample data, inserts a column chart, and sets the chart area to a DarkDownwardDiagonal pattern (light‑gray on white) using Aspose.Cells for .NET, then saves the file as DiagonalStripeChart.xlsx.
class DiagonalStripeChartBackground
{
    static void Main()
    {
        // Create a new workbook
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

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set chart data range
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply diagonal stripe pattern to the chart background
        chart.ChartArea.Area.FillFormat.FillType = FillType.Pattern;
        chart.ChartArea.Area.FillFormat.PatternFill.Pattern = FillPattern.DarkDownwardDiagonal;
        chart.ChartArea.Area.FillFormat.PatternFill.ForegroundColor = Color.LightGray;
        chart.ChartArea.Area.FillFormat.PatternFill.BackgroundColor = Color.White;

        // Save the workbook
        workbook.Save("DiagonalStripeChart.xlsx", SaveFormat.Xlsx);
    }
}
