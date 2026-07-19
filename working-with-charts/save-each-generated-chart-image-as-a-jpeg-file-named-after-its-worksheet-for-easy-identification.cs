// Title: Save each worksheet chart as a JPEG file named after its sheet using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook with three worksheets, adds a column chart to each sheet, and uses Aspose.Cells Chart.ToImage with ImageType.Jpeg to export every chart as a JPEG file whose filename matches the worksheet (e.g., Sheet1.jpg). The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells chart export JPEG | Chart.ToImage C# | save chart image worksheet name | Aspose.Cells multiple chart export | .NET Excel chart to image | dynamic chart filename Aspose
// Common Searches: Aspose.Cells save each chart as JPEG | C# export Excel chart image with worksheet name | Chart.ToImage multiple worksheets Aspose | How to name chart image after sheet in Aspose.Cells | Export Excel charts to JPEG programmatically
// Developer Intent: Export every chart in a workbook to a JPEG file whose name matches its worksheet.
// Use Cases: Generate thumbnail images of sheet‑specific charts for web dashboards. | Create separate chart files for reporting, printing, or email without sharing the full workbook. | Archive visual snapshots of data per sheet for documentation or version control.
// AI Prompts: Write C# code that loops through all worksheets in an Aspose.Cells workbook and saves each chart as a PNG file named after the worksheet. | Show how to add error handling for worksheets that contain no charts when exporting images with Aspose.Cells. | Demonstrate how to set JPEG quality and resolution when using Aspose.Cells Chart.ToImage.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // This C# example creates a workbook with three worksheets, adds a column chart to each sheet, and uses Aspose.Cells Chart.ToImage with ImageType.Jpeg to export every chart as a JPEG file whose filename matches the worksheet (e.g., Sheet1.jpg). The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Ensure the workbook has at least 3 worksheets
                while (workbook.Worksheets.Count < 3)
                {
                    workbook.Worksheets.Add();
                }

                // Add sample data and a chart to each worksheet for demonstration
                for (int wsIndex = 0; wsIndex < 3; wsIndex++)
                {
                    // Get worksheet
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    sheet.Name = $"Sheet{wsIndex + 1}";

                    // Populate some data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["A2"].PutValue("Apple");
                    sheet.Cells["A3"].PutValue("Orange");
                    sheet.Cells["A4"].PutValue("Banana");

                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["B2"].PutValue(10 + wsIndex * 5);
                    sheet.Cells["B3"].PutValue(20 + wsIndex * 5);
                    sheet.Cells["B4"].PutValue(30 + wsIndex * 5);

                    // Add a column chart
                    int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = sheet.Charts[chartIdx];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";

                    // Save the chart image as JPEG named after the worksheet
                    string imagePath = $"{sheet.Name}.jpg";
                    chart.ToImage(imagePath, ImageType.Jpeg);
                    Console.WriteLine($"Chart from '{sheet.Name}' saved to '{imagePath}'.");
                }

                // Save the workbook
                string workbookPath = "ChartsWorkbook.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved as '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
