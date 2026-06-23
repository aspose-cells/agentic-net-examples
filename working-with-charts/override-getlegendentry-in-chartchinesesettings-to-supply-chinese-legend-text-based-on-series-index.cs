using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart.
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");

            sheet.Cells["B1"].PutValue("系列1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            sheet.Cells["C1"].PutValue("系列2");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(40);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories.
            chart.NSeries.Add("B2:C3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Enable the legend.
            chart.ShowLegend = true;

            // Save the workbook.
            string outputPath = "ChartChinese.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}