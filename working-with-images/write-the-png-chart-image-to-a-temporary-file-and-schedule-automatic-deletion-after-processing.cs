// Title: Render an Aspose.Cells column chart to a temporary PNG file and delete it automatically in C#
// AI Prompts: Generate C# code that creates a workbook, adds a column chart, saves the chart as a PNG to a temp file, and guarantees the file is removed after use with Aspose.Cells. | Write a method that uses ImageOrPrintOptions to export an Aspose.Cells chart to a FileStream in the system's temporary directory, then schedules the file for deletion in a finally block.
// Common Searches: c# aspnet save aspose.cells chart as png in temp folder and clean up | how to export a chart to png using Aspose.Cells and delete the temporary file | Aspose.Cells render chart to temporary image file C# example | delete temporary PNG generated from Aspose.Cells chart automatically | using ImageOrPrintOptions to create PNG from Aspose.Cells chart in C#
// Tags: Aspose.Cells chart to PNG export | temporary file cleanup with Aspose.Cells | ImageOrPrintOptions PNG rendering | C# render chart to file stream | auto-delete temp image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Demonstrates creating a workbook, adding sample data and a column chart, exporting the chart to a PNG file placed in the system's temporary directory using ImageOrPrintOptions, and removing the file in a finally block to avoid leftover temporary files.
class ChartToTempPng
{
    static void Main()
    {
        // Path for the temporary PNG file
        string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
        FileInfo tempFileInfo = new FileInfo(tempFilePath);

        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Prepare image options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Optional: set image size
                // ImageHeight = 600,
                // ImageWidth = 800
            };

            // Render the chart directly to a file stream
            using (FileStream fs = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
            {
                chart.ToImage(fs, imgOptions);
            }

            Console.WriteLine("Chart image saved to temporary file: " + tempFilePath);
            // Place any additional processing of the PNG file here
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            // Ensure the temporary file is deleted if it exists
            if (tempFileInfo.Exists)
            {
                try
                {
                    tempFileInfo.Delete();
                    Console.WriteLine("Temporary file deleted.");
                }
                catch (Exception delEx)
                {
                    Console.WriteLine("Failed to delete temporary file: " + delEx.Message);
                }
            }
        }
    }
}
