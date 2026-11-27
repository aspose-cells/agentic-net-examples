// Title: How to hide legends in Aspose.Cells generated Column, Line, and Pie charts and save each as PNG using C#
// Common Searches: C# Aspose.Cells hide chart legend before exporting to PNG | Generate column, line and pie charts without legends using Aspose.Cells .NET | Export multiple Aspose.Cells charts as PNG images for a slideshow in C# | Aspose.Cells ShowLegend property example in C# | Programmatically create chart series and hide legend with Aspose.Cells
// Tags: Aspose.Cells chart ShowLegend false C# | Aspose.Cells export chart to PNG C# | Aspose.Cells generate column line pie charts C# | Aspose.Cells set chart data source C# | Aspose.Cells create multiple charts programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace HideLegendInCharts
{
    // The sample creates a workbook, fills it with sample data, adds Column, Line and Pie charts, disables their legends via ShowLegend = false, assigns titles, exports each chart as a PNG image, and finally saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the charts
                // Columns: Category (A), Series1 (B), Series2 (C), Series3 (D)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["D1"].PutValue("Series 3");

                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
                Random rnd = new Random();

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2
                    sheet.Cells[row, 0].PutValue(categories[i]);               // Column A
                    sheet.Cells[row, 1].PutValue(rnd.Next(10, 100));           // Column B
                    sheet.Cells[row, 2].PutValue(rnd.Next(10, 100));           // Column C
                    sheet.Cells[row, 3].PutValue(rnd.Next(10, 100));           // Column D
                }

                // Define chart types to generate a series of charts for a slideshow
                ChartType[] chartTypes = new ChartType[]
                {
                    ChartType.Column,
                    ChartType.Line,
                    ChartType.Pie
                };

                // Directory to store exported chart images
                string outputDir = Path.Combine(Environment.CurrentDirectory, "ChartImages");
                Directory.CreateDirectory(outputDir);

                // Loop to create each chart, hide its legend, and export as image
                for (int i = 0; i < chartTypes.Length; i++)
                {
                    // Position each chart below the previous one to avoid overlap in the worksheet
                    int upperRow = 2 + i * 20; // Adjust as needed
                    int lowerRow = upperRow + 15;

                    int chartIndex = sheet.Charts.Add(chartTypes[i], upperRow, 0, lowerRow, 8);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set data source for the chart
                    // For Pie chart we only need one series; for others we use all three series
                    if (chartTypes[i] == ChartType.Pie)
                    {
                        chart.NSeries.Add("B2:B7", true);
                        chart.NSeries.CategoryData = "A2:A7";
                    }
                    else
                    {
                        chart.NSeries.Add("B2:D7", true);
                        chart.NSeries.CategoryData = "A2:A7";
                    }

                    // Common chart customizations
                    chart.Title.Text = $"{chartTypes[i]} Chart";
                    chart.Title.Font.IsBold = true;
                    chart.Title.Font.Size = 12;

                    // Hide the legend to reduce visual noise in a slideshow
                    chart.ShowLegend = false;

                    // Save each chart as a separate PNG image
                    string imagePath = Path.Combine(outputDir, $"Chart_{i + 1}_{chartTypes[i]}.png");
                    chart.ToImage(imagePath, ImageType.Png);

                    Console.WriteLine($"Created {chartTypes[i]} chart and saved image to: {imagePath}");
                }

                // Save the workbook containing all charts
                string workbookPath = Path.Combine(Environment.CurrentDirectory, "ChartsWorkbook.xlsx");
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {workbookPath}");

                Console.WriteLine("All charts generated with legends hidden. Press any key to exit.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
