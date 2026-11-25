// Title: Setting a chart title to "Monthly Revenue" and applying an Accent1 font color with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart title text and make it visible | apply Accent1 color to chart title font using Aspose.Cells .NET | change column chart title color to theme accent in Aspose.Cells | customize chart title appearance in an Excel workbook with Aspose.Cells C# | save workbook with styled chart title using Aspose.Cells for .NET
// Tags: Aspose.Cells set chart title visibility | Aspose.Cells apply Accent1 font color to chart title | Aspose.Cells column chart title styling | Aspose.Cells save workbook with styled chart | Aspose.Cells .NET chart title customization

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds month and revenue data, inserts a column chart, sets the chart title to "Monthly Revenue" with visibility enabled, applies an Accent1-like color to the title font, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B2"].PutValue(12000);
            worksheet.Cells["B3"].PutValue(15000);
            worksheet.Cells["B4"].PutValue(18000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title
            chart.Title.Text = "Monthly Revenue";
            chart.Title.IsVisible = true;

            // Apply a predefined Accent1-like color to the title font
            // (Using a direct color as a fallback when BuiltInStyle is unavailable)
            chart.Title.Font.Color = Color.FromArgb(91, 155, 213); // Approximate Accent1 color

            // Save the workbook
            string outputPath = "MonthlyRevenueChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
