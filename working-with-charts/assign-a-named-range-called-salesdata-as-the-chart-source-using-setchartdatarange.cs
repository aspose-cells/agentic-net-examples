using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

class AssignNamedRangeToChart
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart (A1:B4)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Create a named range called "SalesData" that refers to the data area
            AsposeRange dataRange = sheet.Cells.CreateRange("A1", "B4");
            dataRange.Name = "SalesData";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Assign the named range as the chart source (plotted by column)
            chart.SetChartDataRange("SalesData", true);

            // Save the workbook
            string outputPath = "ChartWithNamedRange.xlsx";
            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}