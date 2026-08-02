using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SetChartTitleWithAccentStyle
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["B4"].PutValue(18000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title text and make it visible
            chart.Title.Text = "Monthly Revenue";
            chart.Title.IsVisible = true;

            // Apply the Accent1 theme color to the title font
            Color accent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
            chart.Title.Font.Color = accent1;

            // Save the workbook
            string outputPath = "ChartWithTitleAccent1.xlsx";

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}