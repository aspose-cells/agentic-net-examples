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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data for the chart
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Quarter");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);

            // Add a column chart that uses the data above
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Categories

            // Ensure the output folder exists
            string outputFolder = "output";
            Directory.CreateDirectory(outputFolder);

            // Save the workbook as an XLSX file named QuarterlyReport.xlsx
            string outputPath = Path.Combine(outputFolder, "QuarterlyReport.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}