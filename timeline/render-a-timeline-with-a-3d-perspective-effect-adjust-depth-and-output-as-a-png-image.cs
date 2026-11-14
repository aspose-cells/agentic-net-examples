// Title: Generate a 3‑D column chart timeline and save it as a PNG image using Aspose.Cells for .NET (C#)
// Common Searches: how to create a 3d column chart from date data in C# with Aspose.Cells and export to PNG | Aspose.Cells C# render a timeline chart with depth effect and save as image file | export Excel 3d column chart to PNG using Aspose.Cells .NET API | C# code to add a 3d column chart to a workbook and convert it to PNG image | save Aspose.Cells chart as PNG with custom title and data range
// Tags: Aspose.Cells 3D column chart generation | C# render chart as PNG image | timeline chart depth effect Aspose.Cells | chart data source setup Aspose.Cells | save workbook chart to image file .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, fills columns with dates and values, adds a 3‑D column chart to represent a timeline, sets the chart title, and saves the workbook as a PNG file, which renders only the chart as an image.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data for the timeline (dates in column A, values in column B)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
                sheet.Cells["B4"].PutValue(15);
                sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
                sheet.Cells["B5"].PutValue(25);
                sheet.Cells["A6"].PutValue(new DateTime(2023, 5, 1));
                sheet.Cells["B6"].PutValue(30);

                // Add a 3‑D column chart (Aspose.Cells does not have a Timeline chart type)
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 7, 0, 27, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Optional: set a title
                chart.Title.Text = "Project Timeline (3‑D)";

                // Define output file path
                string outputPath = "timeline.png";

                // Save the workbook as a PNG image (the chart will be rendered to the image)
                workbook.Save(outputPath, SaveFormat.Png);

                Console.WriteLine($"Chart image saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
