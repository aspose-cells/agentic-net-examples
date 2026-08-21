// Title: Toggle picture visibility on an Aspose.Cells chart with a Boolean flag (C#)
// Description: Creates a new workbook, adds a column chart with sample data, and uses a Boolean variable to decide whether to insert a picture (example.png) near the chart. The picture is positioned by cell coordinates and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# toggle picture | conditional image insertion Excel chart | hide/show picture Aspose.Cells | chart overlay image visibility .NET | boolean flag picture Aspose.Cells
// Common Searches: Aspose.Cells show picture on chart conditionally | C# add picture to Excel chart only if flag is true | toggle visibility of chart image Aspose.Cells | conditional picture insertion Excel using Aspose | how to hide a picture in an Aspose.Cells chart
// Developer Intent: Insert or omit a picture on an Excel chart at runtime based on a Boolean condition using Aspose.Cells for .NET.
// Use Cases: Add a company logo to reports only when branding is required. | Display a warning icon on a chart when data exceeds thresholds. | Show a trend‑line graphic when a specific analysis mode is enabled.
// AI Prompts: Write C# code that adds a picture to an Aspose.Cells chart only when a boolean variable is true. | Show how to remove or hide an existing picture from an Aspose.Cells chart when a condition changes. | Explain positioning a picture relative to a chart and controlling its visibility with Aspose.Cells APIs.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds a column chart with sample data, and uses a Boolean variable to decide whether to insert a picture (example.png) near the chart. The picture is positioned by cell coordinates and the workbook is saved as an XLSX file.
    public class TogglePictureVisibilityInChart
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
            // Flag that determines whether the picture should be visible
            bool showPicture = true; // set to false to hide the picture

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a sample chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Populate chart data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a picture to the worksheet if required and the file exists
            string imagePath = "example.png";

            if (showPicture)
            {
                if (File.Exists(imagePath))
                {
                    try
                    {
                        int pictureIndex = worksheet.Pictures.Add(6, 1, imagePath);
                        Picture picture = worksheet.Pictures[pictureIndex];

                        // Position the picture near the chart (using fixed cells as reference)
                        picture.UpperLeftRow = 6;      // row index (zero‑based)
                        picture.UpperLeftColumn = 1;   // column index (zero‑based)
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to add picture: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture addition.");
                }
            }

            // Save the workbook
            try
            {
                workbook.Save("TogglePictureVisibilityInChart.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
