using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SaveChartWorkbook
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Specify the output directory and file name
        string outputDirectory = @"C:\Output";
        Directory.CreateDirectory(outputDirectory); // Ensure the directory exists
        string outputPath = Path.Combine(outputDirectory, "ChartWorkbook.xlsx");

        // Save the workbook (which includes the chart) as an XLSX file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}