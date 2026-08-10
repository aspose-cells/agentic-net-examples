// Title: Save an Aspose.Cells workbook as a PNG with a timestamped filename in C#
// Description: Creates a workbook, adds sample data and a column chart, ensures a target folder exists, builds a filename that includes the current date and time, configures PNG rendering options, and uses WorkbookRender to export the first worksheet as a PNG image saved to the generated path.
// Keywords: Aspose.Cells PNG export C# | WorkbookRender timestamped file | ImageOrPrintOptions SaveFormat.Png | C# create output directory | Excel to image Aspose.Cells | timestamped filename .NET | render chart as PNG Aspose | save worksheet as image C#
// Common Searches: Aspose.Cells export worksheet to PNG with date in filename | C# save rendered Excel chart as PNG in specific folder | How to add timestamp to Aspose.Cells image file name | Create folder and save PNG using Aspose.Cells .NET | WorkbookRender PNG output example
// Developer Intent: Generate a PNG snapshot of a workbook (or its chart) and store it in a designated directory using a filename that contains the current timestamp.
// Use Cases: Daily automated reports that archive the latest worksheet view as a dated PNG image. | Batch processing pipelines that convert multiple workbooks to uniquely named PNG files for audit trails. | Web services that return chart images, naming each file with a timestamp to prevent naming collisions.
// AI Prompts: Write C# code with Aspose.Cells to render the first worksheet to a PNG file in a given folder, appending the current date and time to the filename. | Provide a reusable method that accepts a Workbook and an output path, creates the directory if missing, and saves the workbook as a PNG with a timestamp while handling errors. | Explain how to set ImageOrPrintOptions for PNG output and use WorkbookRender to produce a timestamped image file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data and a column chart, ensures a target folder exists, builds a filename that includes the current date and time, configures PNG rendering options, and uses WorkbookRender to export the first worksheet as a PNG image saved to the generated path.
class SaveWorkbookAsPngWithTimestamp
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart (optional, demonstrates content)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.SetChartDataRange("A1:B4", true);

            // Define output directory and ensure it exists
            string outputDir = "GeneratedImages";
            Directory.CreateDirectory(outputDir);

            // Build timestamped filename
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string filePath = Path.Combine(outputDir, $"Workbook_{timestamp}.png");

            // Configure PNG rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            imgOptions.SaveFormat = SaveFormat.Png;

            // Render the first worksheet to PNG
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
            renderer.ToImage(0, filePath);

            Console.WriteLine("Workbook rendered and saved as PNG to: " + filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
