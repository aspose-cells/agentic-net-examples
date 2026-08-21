// Title: C# – Insert a picture into an Excel chart and set its Z‑order behind the chart shape using Aspose.Cells
// Description: Creates a new workbook, adds a column chart, reads an image file, inserts the picture into the chart with cell‑based coordinates via AddPictureInChart, moves the picture one layer back with ToFrontOrBack(-1) so it stays behind the chart shape, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | Insert picture into chart | AddPictureInChart | Z‑order | ToFrontOrBack | chart shape layering | Excel chart image | save workbook
// Common Searches: Aspose.Cells insert image into chart C# | How to change Z‑order of shapes in Excel chart using Aspose.Cells | AddPictureInChart example .NET | Move picture behind chart series Aspose.Cells | Set picture layer back in Excel chart programmatically
// Developer Intent: Add an image to a chart, place it behind the chart elements by adjusting its Z‑order, and save the workbook.
// Use Cases: Add a company logo to a chart while keeping data series visible on top. | Apply a watermark to a chart without obscuring the plotted data. | Control the layering of multiple shapes in a chart for custom reporting layouts.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a PNG into a line chart and positions it behind the chart series. | Explain the ToFrontOrBack method for chart shapes in Aspose.Cells and show how to move a picture multiple layers back. | Provide error‑handling best practices for missing image files when adding pictures to charts with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a new workbook, adds a column chart, reads an image file, inserts the picture into the chart with cell‑based coordinates via AddPictureInChart, moves the picture one layer back with ToFrontOrBack(-1) so it stays behind the chart shape, and saves the file as an XLSX document.
    class InsertPictureWithZOrder
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a simple column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("A1:A3", true);
                chart.NSeries.CategoryData = "B1:B3";

                // Path to the image file
                string imagePath = "example.jpg";

                // Insert picture into the chart if the image file exists
                if (File.Exists(imagePath))
                {
                    byte[] imageBytes = File.ReadAllBytes(imagePath);
                    using (MemoryStream imageStream = new MemoryStream(imageBytes))
                    {
                        // Add picture to the chart (position defined by cell coordinates)
                        // Correct parameter order: upperLeftRow, upperLeftColumn, pictureStream, lowerRightRow, lowerRightColumn
                        Picture picture = chart.Shapes.AddPictureInChart(0, 0, imageStream, 10, 10);

                        // Send the picture one position back so it stays behind the chart shape
                        picture.ToFrontOrBack(-1);
                    }
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                workbook.Save("InsertPictureWithZOrder.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
