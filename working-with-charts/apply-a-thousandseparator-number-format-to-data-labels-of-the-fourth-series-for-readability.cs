using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ApplyThousandSeparatorToFourthSeries
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for four series
                // Categories (X‑axis)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["B4"].PutValue(1800);
                sheet.Cells["B5"].PutValue(2100);

                // Series 2
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(2200);
                sheet.Cells["C3"].PutValue(2500);
                sheet.Cells["C4"].PutValue(2800);
                sheet.Cells["C5"].PutValue(3100);

                // Series 3
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(3200);
                sheet.Cells["D3"].PutValue(3500);
                sheet.Cells["D4"].PutValue(3800);
                sheet.Cells["D5"].PutValue(4100);

                // Series 4 (the target series)
                sheet.Cells["E1"].PutValue("Series4");
                sheet.Cells["E2"].PutValue(4200);
                sheet.Cells["E3"].PutValue(4500);
                sheet.Cells["E4"].PutValue(4800);
                sheet.Cells["E5"].PutValue(5100);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add all four series to the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Series1";

                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Series2";

                chart.NSeries.Add("D2:D5", true);
                chart.NSeries[2].Name = "Series3";

                chart.NSeries.Add("E2:E5", true);
                chart.NSeries[3].Name = "Series4";

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the fourth series (index 3)
                Series fourthSeries = chart.NSeries[3];
                fourthSeries.DataLabels.ShowValue = true;

                // Apply thousand‑separator number format to the data labels
                fourthSeries.DataLabels.NumberFormat = "#,##0";

                // Save the workbook
                string outputPath = "ThousandSeparatorFourthSeries.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyThousandSeparatorToFourthSeries.Run();
        }
    }
}