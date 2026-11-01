// Title: Insert a picture from a FileStream into an Excel worksheet and set its width to 100 points using Aspose.Cells for .NET
// AI Prompts: Read a PNG file into a FileStream, add it as a picture on a worksheet, assign the Width property a value of 100 points, and save the workbook with Aspose.Cells in C#. | Create a column chart, overlay an image loaded from a stream as a picture, adjust the picture's width to 100 points, and export the file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add picture from FileStream to worksheet and set width | how to set picture width in points using Aspose.Cells | load PNG into stream and insert into Excel workbook with Aspose.Cells .NET | resize picture inserted via Aspose.Cells C# example | place image over chart in Aspose.Cells worksheet
// Tags: image stream picture insertion Aspose.Cells | picture width 100 points Aspose.Cells | PNG overlay on worksheet chart Aspose.Cells | worksheet picture sizing Aspose.Cells | Aspose.Cells picture control dimensions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// // This example creates a workbook, adds a column chart, loads a PNG image via FileStream, inserts it as a picture on the worksheet, sets the picture width to 100 points, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (position and size are arbitrary)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Path to the image file
            string imagePath = "image.png";

            // Verify that the image file exists before loading
            if (File.Exists(imagePath))
            {
                try
                {
                    // Load the image file into a stream
                    using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        // Add a picture to the worksheet; Add returns the picture index
                        int pictureIndex = sheet.Pictures.Add(0, 0, 0, 0, imgStream);
                        Picture picture = sheet.Pictures[pictureIndex];

                        // Set the picture width to 100 points
                        picture.Width = 100;
                    }
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to add picture: {imgEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}");
            }

            // Save the workbook to a file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
