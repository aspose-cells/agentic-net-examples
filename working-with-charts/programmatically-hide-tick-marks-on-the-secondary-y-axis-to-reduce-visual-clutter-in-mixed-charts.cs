// Title: Programmatically hide tick marks on the secondary Y‑axis of a mixed column‑line chart with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide secondary axis tick marks in mixed chart | remove major and minor tick marks from secondary value axis using Aspose.Cells | how to suppress secondary Y axis tick marks in Excel chart via Aspose.Cells API | C# Aspose.Cells set secondary axis MajorTickMark to None | customize secondary axis appearance in mixed column line chart Aspose.Cells
// Tags: secondary value axis tick mark removal Aspose.Cells | mixed column line chart secondary axis styling Aspose.Cells | Chart.SecondValueAxis MajorTickMark None C# example | Aspose.Cells programmatic axis tick suppression | Excel chart secondary axis customization using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sales and profit data, builds a mixed column‑line chart, plots the profit series on the secondary Y‑axis, and programmatically sets both MajorTickMark and MinorTickMark of the secondary value axis to TickMarkType.None while keeping the axis visible and titled, then saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a mixed chart (Column + Line)
        // Categories
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        // Primary series (Column) - Sales
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["B5"].PutValue(200);

        // Secondary series (Line) - Profit
        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(55);
        sheet.Cells["C5"].PutValue(70);

        // Add a mixed chart (Column for primary series, Line for secondary series)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add primary series (Column) and bind it to the first data range
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries[0].Name = "Sales";

        // Add secondary series (Line) and bind it to the second data range
        chart.NSeries.Add("C2:C5", true);
        chart.NSeries[1].Name = "Profit";

        // Set the category (X) axis data
        chart.NSeries.CategoryData = "A2:A5";

        // Plot the second series on the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Change the chart type of the second series to Line
        chart.NSeries[1].Type = ChartType.Line;

        // Hide tick marks on the secondary Y axis
        Axis secondaryValueAxis = chart.SecondValueAxis;
        secondaryValueAxis.MajorTickMark = TickMarkType.None;
        secondaryValueAxis.MinorTickMark = TickMarkType.None;

        // (Optional) Keep the secondary axis visible so the data is still readable
        secondaryValueAxis.IsVisible = true;
        secondaryValueAxis.Title.Text = "Profit (Secondary Axis)";

        // Save the workbook
        string outputPath = "MixedChartSecondaryAxisTickMarksHidden.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
