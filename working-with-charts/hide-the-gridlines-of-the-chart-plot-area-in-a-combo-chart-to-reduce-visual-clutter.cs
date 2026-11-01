// Title: Hide major gridlines on both value and category axes of a combo chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a column‑line combo chart and disables the major gridlines on the value and category axes. | Show how to programmatically turn off plot area gridlines for a combo chart in an Excel workbook with Aspose.Cells for .NET. | Provide a snippet that adds sample data, builds a combo chart, and sets chart.ValueAxis.MajorGridLines.IsVisible = false and chart.CategoryAxis.MajorGridLines.IsVisible = false before saving.
// Common Searches: Aspose.Cells C# hide chart gridlines combo chart | disable major gridlines on Excel chart axes using Aspose.Cells .NET | remove plot area gridlines from column line combo chart programmatically | how to turn off value axis gridlines in Aspose.Cells generated workbook | C# Aspose.Cells chart formatting hide category axis gridlines
// Tags: Aspose.Cells hide chart gridlines C# | combo chart major gridlines visibility Aspose.Cells | disable value axis gridlines Aspose.Cells .NET | remove category axis gridlines Aspose.Cells | Excel chart formatting Aspose.Cells C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Creates a workbook, adds sample data, builds a column‑line combo chart, and hides both value and category major gridlines before saving the file as ComboChart_NoGridlines.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);

            // Add a combo chart (column chart as base) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sales and Profit";

            // First series as Column (Sales)
            int series1 = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[series1].Name = "Sales";
            chart.NSeries[series1].Type = ChartType.Column;

            // Second series as Line (Profit)
            int series2 = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[series2].Name = "Profit";
            chart.NSeries[series2].Type = ChartType.Line;

            // Hide major gridlines in the plot area (value and category axes)
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.CategoryAxis.MajorGridLines.IsVisible = false;

            // Define output file path
            string outputPath = "ComboChart_NoGridlines.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
