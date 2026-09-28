// Title: How to hide major gridlines in a column‑line combo chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that creates an Excel workbook, adds sample data, builds a combo chart with a column series and a line series, and disables the major gridlines on both the value and category axes using Aspose.Cells. | Generate code that constructs a combo chart with Aspose.Cells, then sets chart.ValueAxis.MajorGridLines.IsVisible and chart.CategoryAxis.MajorGridLines.IsVisible to false before saving the workbook.
// Common Searches: asp.net aspose.cells hide major gridlines on combo chart | c# generate Excel combo chart without axis gridlines using Aspose | remove plot area gridlines from column and line chart programmatically Aspose.Cells | how to disable chart gridlines in Aspose.Cells C# example | Aspose.Cells combo chart formatting hide gridlines
// Tags: Aspose.Cells hide chart gridlines | combo chart without gridlines C# | disable major gridlines chart axis Aspose | Excel chart formatting Aspose.Cells | column line combo chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates it with sample data, adds a combo chart that combines a column series and a line series, assigns category labels, hides the major gridlines on both the value and category axes, and saves the file as ComboChart_NoGridlines.xlsx.
class HideChartGridlines
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart (used as the base for a combo chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add first series (Column)
            int seriesIndex1 = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[seriesIndex1].Name = "Series 1";

            // Add second series (Line)
            int seriesIndex2 = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[seriesIndex2].Name = "Series 2";
            chart.NSeries[seriesIndex2].Type = ChartType.Line;

            // Assign categories
            chart.NSeries.CategoryData = "A2:A4";

            // Hide major gridlines on both axes
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.CategoryAxis.MajorGridLines.IsVisible = false;

            // Save the workbook
            string outputPath = "ComboChart_NoGridlines.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
