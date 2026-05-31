using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Quarter");
        worksheet.Cells["A2"].PutValue("Q1");
        worksheet.Cells["A3"].PutValue("Q2");
        worksheet.Cells["A4"].PutValue("Q3");
        worksheet.Cells["A5"].PutValue("Q4");

        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(130);
        worksheet.Cells["B5"].PutValue(170);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);               // Values
        chart.NSeries.CategoryData = "A2:A5";           // Categories

        // Ensure the output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Save the workbook containing the chart as an XLSX file
        string filePath = Path.Combine(outputDir, "QuarterlyReport.xlsx");
        workbook.Save(filePath, SaveFormat.Xlsx);
    }
}