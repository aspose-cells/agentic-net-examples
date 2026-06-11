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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Fill worksheet with sample data -----
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");

            // First series values (column B)
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Second series values (column C)
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // ----- Add a column chart -----
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add multiple series
            chart.NSeries.Add("=Sheet1!$B$2:$B$5", true); // Series 1
            chart.NSeries.Add("=Sheet1!$C$2:$C$5", true); // Series 2

            // Set the category (X‑axis) data for the chart
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

            // ----- Freeze rows that contain the series data -----
            // Freeze rows 1‑5 (row index is 0‑based, so split at row 5)
            // Provide total rows/columns to define the frozen area
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalCols = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(5, 0, totalRows, totalCols);

            // Ensure output directory exists
            string outputPath = "MultipleSeriesWithFreeze.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}