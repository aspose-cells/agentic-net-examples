// Title: Export an Aspose.Cells column chart to a PNG file via a MemoryStream in C#
// Common Searches: C# Aspose.Cells export chart to PNG using MemoryStream | How to save a chart image from Aspose.Cells to disk in .NET | Write Aspose.Cells chart image to file without creating a temporary file | Using ImageOrPrintOptions to render an Aspose.Cells chart as PNG | Aspose.Cells chart.ToImage example with stream handling
// Tags: Aspose.Cells chart export PNG | chart.ToImage MemoryStream C# | ImageOrPrintOptions chart rendering | write chart image to file .NET | Aspose.Cells column chart image generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The code builds a workbook, adds a column chart, renders the chart to a PNG via ImageOrPrintOptions and Chart.ToImage, stores the PNG in a MemoryStream, resets the stream position, creates the target directory if needed, and copies the stream to a FileStream to save the image file.
class ExportChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set image export options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Export the chart to a PNG file via a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                chart.ToImage(ms, imgOptions);
                ms.Position = 0; // Reset stream position before writing

                string outputPath = "ExportedChart.png";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write the image to file
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(file);
                }

                Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
            }

            // Optional: Save the workbook if needed
            // workbook.Save("WorkbookWithChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
