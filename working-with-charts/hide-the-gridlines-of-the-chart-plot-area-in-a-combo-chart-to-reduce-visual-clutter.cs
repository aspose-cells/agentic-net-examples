// Title: Hide major and minor gridlines in a combo chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide major gridlines in a combo chart | remove minor gridlines from Excel combo chart with Aspose.Cells .NET | how to disable both major and minor chart axis gridlines using Aspose.Cells | Aspose.Cells chart formatting hide all gridlines in combo chart | reduce visual clutter in Excel combo chart by hiding gridlines C#
// Tags: Aspose.Cells combo chart axis gridlines control | chart gridlines visibility settings .NET | Excel chart formatting without gridlines C# | customize chart appearance Aspose.Cells | remove chart axis lines Aspose.Cells API | visual clutter reduction in Excel charts

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, builds a combo chart, and then disables both major and minor gridlines on the value and category axes before saving the file, resulting in a cleaner chart layout.
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

            // Add a chart (using Column as base type; series types can be changed individually if needed)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Series 1";
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Series 2";

            // Set the category axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Hide major and minor gridlines on both axes
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.CategoryAxis.MajorGridLines.IsVisible = false;
            chart.ValueAxis.MinorGridLines.IsVisible = false;
            chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Define output file path
            string outputPath = "ComboChart_NoGridlines.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
