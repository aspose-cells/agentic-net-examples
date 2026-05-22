using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendNegativeFill
{
    class Program
    {
        static void Main()
        {
            try
            {
                Workbook workbook;

                // If a template file is needed, load it safely; otherwise create a new workbook
                string templatePath = "Template.xlsx";
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with both positive and negative values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(50);   // positive
                sheet.Cells["B3"].PutValue(-30);  // negative

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(20);   // positive
                sheet.Cells["C3"].PutValue(40);   // positive

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:C3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Customize fill colors for positive and negative points
                foreach (Series series in chart.NSeries)
                {
                    // Positive points – blue fill
                    series.Area.ForegroundColor = Color.Blue;
                    // Negative points – red fill
                    series.Area.BackgroundColor = Color.Red;
                }

                // Save the workbook
                string outputPath = "LegendNegativeFill.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}