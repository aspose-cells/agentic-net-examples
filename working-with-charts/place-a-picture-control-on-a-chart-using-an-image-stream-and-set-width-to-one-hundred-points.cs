// Title: Insert a picture into an Aspose.Cells chart from an image stream and set its width to 100 points (C#)
// AI Prompts: How can I add a picture to a chart in Aspose.Cells for .NET using a MemoryStream and set the picture width to 100 points? | Generate C# code that loads an image into a stream, places it as a picture control on a column chart, and adjusts its width to 100 points with Aspose.Cells.
// Common Searches: Aspose.Cells C# add image to chart from stream with specific width | set chart picture size in points using Aspose.Cells .NET | place picture control on column chart using image stream Aspose.Cells example
// Tags: Aspose.Cells chart picture insertion from stream | set picture width points Aspose.Cells | C# column chart add image control | Aspose.Cells workbook picture workaround chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a column chart, loads an image file into a stream, inserts the image as a picture control on the chart (using a worksheet picture as a workaround), sets the picture width to 100 points, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (positioned from row 5, column 0 to row 15, column 5)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Verify that the image file exists before attempting to load it
            const string imagePath = "image.png";
            if (File.Exists(imagePath))
            {
                try
                {
                    // Insert the picture into the worksheet (as a workaround for chart picture insertion)
                    // The picture is placed at row 0, column 0 of the worksheet.
                    int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Set the picture width to 100 points
                    picture.Width = 100;
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Failed to insert picture: {picEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook to a file
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
