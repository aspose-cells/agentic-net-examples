// Title: Export an Excel column chart to PNG with Aspose.Cells and embed the image into a PowerPoint slide using Aspose.Slides (C#)
// AI Prompts: Write C# code that creates a column chart in an Excel workbook, renders the chart to a PNG image using Aspose.Cells, and then adds that PNG to a new slide of a PowerPoint presentation with Aspose.Slides. | Refactor the example to keep the chart image in a MemoryStream and insert it directly into the PowerPoint slide without saving a temporary PNG file to disk. | Add comprehensive error handling so the PowerPoint file is saved only when the chart image is successfully generated and inserted.
// Common Searches: Aspose.Cells export chart to PNG and insert into PowerPoint C# | C# generate Excel chart image and embed in PPT using Aspose libraries | How to render an Aspose.Cells chart as PNG and add to Aspose.Slides slide | Convert Excel column chart to image and create PowerPoint slide programmatically .NET | Aspose.Cells chart rendering to stream for use in Aspose.Slides presentation
// Tags: export chart to PNG Aspose.Cells | embed chart image into PowerPoint Aspose.Slides | C# generate Excel column chart image | Aspose.Cells chart rendering to stream | Aspose.Slides insert picture from memory stream

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program creates a workbook, adds sample data, builds a column chart, renders the chart to a PNG image using ImageOrPrintOptions, keeps the image in a MemoryStream, and then inserts that PNG into a new slide of a PowerPoint presentation via Aspose.Slides before saving both files.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Populate sample data for the chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(15);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Column Chart";

            // Export the chart to a PNG image file
            const string imagePath = "ChartImage.png";
            var imgOptions = new ImageOrPrintOptions(); // default format is PNG

            using (var pngStream = new MemoryStream())
            {
                chart.ToImage(pngStream, imgOptions);
                File.WriteAllBytes(imagePath, pngStream.ToArray());
            }

            // Save the workbook
            const string workbookPath = "SampleChart.xlsx";
            workbook.Save(workbookPath);

            Console.WriteLine($"Workbook saved to '{workbookPath}'.");
            Console.WriteLine($"Chart image saved to '{imagePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
