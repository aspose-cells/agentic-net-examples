using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetXAxisTitleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (years and sales)
                sheet.Cells["A1"].PutValue("Year");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue(2018);
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["A3"].PutValue(2019);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue(2020);
                sheet.Cells["B4"].PutValue(250);
                sheet.Cells["A5"].PutValue(2021);
                sheet.Cells["B5"].PutValue(300);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // X‑axis categories (years)

                // Set a descriptive title for the X‑axis
                chart.CategoryAxis.Title.Text = "Year (2018‑2021)";
                chart.CategoryAxis.Title.IsVisible = true;
                chart.CategoryAxis.Title.Font.IsBold = true; // optional bold

                // Save the workbook
                string outputPath = "SetXAxisTitleDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetXAxisTitleDemo.Run();
        }
    }
}