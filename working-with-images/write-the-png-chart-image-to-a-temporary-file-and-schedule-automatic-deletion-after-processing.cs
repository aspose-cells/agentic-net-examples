// Title: Render an Aspose.Cells chart to a temporary PNG file and auto‑delete it on application exit (C#)
// AI Prompts: Write C# code that creates a workbook chart with Aspose.Cells, saves it as a PNG in the system temp folder, and registers a ProcessExit handler to remove the file. | Show how to use ImageOrPrintOptions to export an Aspose.Cells chart to a temporary PNG image and ensure the file is deleted when the program terminates.
// Common Searches: how to save an Aspose.Cells chart as a PNG in the temp directory using C# | auto delete temporary chart image on program exit Aspose.Cells | C# Aspose.Cells render chart to image and clean up temporary file | using ImageOrPrintOptions to export chart to PNG and schedule deletion | temporary file handling for Aspose.Cells chart images in .NET
// Tags: Aspose.Cells chart to PNG export | render chart to temporary PNG file | register ProcessExit for temp file deletion C# | ImageOrPrintOptions chart rendering | auto‑delete temporary chart image .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, populates data, adds a column chart, and uses Aspose.Cells ImageOrPrintOptions to render the chart directly to a PNG file in the system's temporary folder. It then registers a ProcessExit event that deletes the temporary file when the application ends, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Prepare image options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Generate a temporary file path for the chart image
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

            // Render the chart directly to the temporary file
            chart.ToImage(tempFilePath, imgOptions);

            // Ensure the temporary file is deleted when the process exits
            AppDomain.CurrentDomain.ProcessExit += (sender, args) =>
            {
                try
                {
                    if (File.Exists(tempFilePath))
                        File.Delete(tempFilePath);
                }
                catch
                {
                    // Optionally log deletion errors
                }
            };

            Console.WriteLine("Chart image saved to temporary file: " + tempFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
