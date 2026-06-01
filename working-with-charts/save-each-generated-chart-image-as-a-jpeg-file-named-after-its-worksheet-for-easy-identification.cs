using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    public class ExportChartsAsJpeg
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Populate worksheets with sample data and charts
                for (int i = 0; i < 3; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    sheet.Name = $"Sheet{i + 1}";

                    // Sample data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["A2"].PutValue("Apple");
                    sheet.Cells["A3"].PutValue("Orange");
                    sheet.Cells["A4"].PutValue("Banana");

                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["B2"].PutValue(10 + i * 5);
                    sheet.Cells["B3"].PutValue(15 + i * 5);
                    sheet.Cells["B4"].PutValue(7 + i * 5);

                    // Add a column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";
                }

                // Export each chart as a JPEG file named after its worksheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Chart chart in ws.Charts)
                    {
                        try
                        {
                            string imagePath = $"{ws.Name}.jpg";
                            chart.ToImage(imagePath, ImageType.Jpeg);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error exporting chart from worksheet '{ws.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the workbook
                workbook.Save("WorkbookWithCharts.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}