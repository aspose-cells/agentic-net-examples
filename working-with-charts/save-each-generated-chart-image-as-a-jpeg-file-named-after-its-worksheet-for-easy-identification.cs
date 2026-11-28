// Title: Export charts from each worksheet to JPEG files named after the worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export worksheet chart as JPEG with worksheet name | How to save Excel charts to separate JPEG files per sheet using Aspose.Cells | Aspose.Cells chart.ToImage JPEG filename based on worksheet name | Generate JPEG images for all charts in a workbook with safe file names in .NET | Batch export multiple charts to JPEG in Aspose.Cells C# loop through worksheets
// Tags: chart to jpeg export Aspose.Cells | worksheet based chart image naming | safe filename generation for chart export | batch export charts from workbook .NET | using ImageType.Jpeg Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartImageExporter
{
    // The example creates a workbook with two worksheets, adds sample data and charts, then creates a "ChartImages" folder. It iterates through each worksheet, builds a filesystem‑safe filename that includes the worksheet name and chart index, and saves each chart as a JPEG using chart.ToImage with ImageType.Jpeg.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output folder where JPEG files will be saved
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ChartImages");
            Directory.CreateDirectory(outputFolder);

            // Create a new workbook and add sample worksheets with charts
            Workbook workbook = new Workbook();

            // Worksheet 1
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "SalesData";
            PopulateSampleData(ws1);
            AddSampleChart(ws1, ChartType.Column, 5, 0, 20, 8);

            // Worksheet 2
            Worksheet ws2 = workbook.Worksheets.Add("ProfitData");
            PopulateSampleData(ws2);
            AddSampleChart(ws2, ChartType.Pie, 5, 0, 20, 8);

            // Save the workbook (optional, just to keep a copy)
            string workbookPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleWorkbook.xlsx");
            workbook.Save(workbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {workbookPath}");

            // Iterate through each worksheet and export its charts as JPEG images
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains no charts.");
                    continue;
                }

                for (int i = 0; i < sheet.Charts.Count; i++)
                {
                    Chart chart = sheet.Charts[i];

                    // Build a file name that includes the worksheet name and chart index
                    string safeSheetName = MakeFileSystemSafe(sheet.Name);
                    string imageFileName = $"{safeSheetName}_Chart{i + 1}.jpg";
                    string imagePath = Path.Combine(outputFolder, imageFileName);

                    try
                    {
                        // Export chart to JPEG using the overload that accepts ImageType
                        chart.ToImage(imagePath, ImageType.Jpeg);
                        Console.WriteLine($"Chart exported: {imagePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart from worksheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Processing completed. Press any key to exit.");
            Console.ReadKey();
        }

        private static void PopulateSampleData(Worksheet sheet)
        {
            // Header
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Sample rows
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(800);

            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(1500);
        }

        /// <summary>
        /// Adds a chart to the specified worksheet using the provided range.
        /// </summary>
        private static void AddSampleChart(Worksheet sheet, ChartType type, int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn)
        {
            int chartIndex = sheet.Charts.Add(type, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);
        }

        /// <summary>
        /// Replaces characters that are invalid in file names.
        /// </summary>
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
