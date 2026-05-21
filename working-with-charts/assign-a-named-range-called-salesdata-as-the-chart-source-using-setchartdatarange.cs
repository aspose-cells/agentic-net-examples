using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetChartDataRangeWithNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(250);
                sheet.Cells["A5"].PutValue("Q4");
                sheet.Cells["B5"].PutValue(300);

                // Create a named range called "SalesData" that refers to the data area A1:B5
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                // RefersTo must start with '=' and include the sheet name
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$5";

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart's data source to the named range "SalesData"
                // The second parameter (true) indicates that data is plotted by column
                chart.SetChartDataRange("SalesData", true);

                // Optionally set a title for clarity
                chart.Title.Text = "Quarterly Sales";

                // Determine output path and ensure directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithNamedRange.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during workbook creation: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetChartDataRangeWithNamedRange.Run();
        }
    }
}