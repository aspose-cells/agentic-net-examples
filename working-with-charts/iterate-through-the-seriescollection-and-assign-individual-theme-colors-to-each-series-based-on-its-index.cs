using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class SeriesCollectionThemeColorsDemo
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 12);
                sheet.Cells[$"D{i}"].PutValue(i * 14);
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart data range and category data
            chart.NSeries.Add("B2:D6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Define a list of theme colors to apply (Accent1, Accent2, Accent3, etc.)
            ThemeColorType[] themeColors = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Apply theme colors to each series
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                ThemeColorType colorType = themeColors[i % themeColors.Length];
                chart.NSeries[i].Border.ThemeColor = new ThemeColor(colorType, 0.0);
                chart.NSeries[i].Border.Style = LineType.Solid;
                chart.NSeries[i].Border.Weight = WeightType.MediumLine;
            }

            // Ensure output directory exists
            string outputPath = "SeriesCollection_ThemeColors_Demo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}