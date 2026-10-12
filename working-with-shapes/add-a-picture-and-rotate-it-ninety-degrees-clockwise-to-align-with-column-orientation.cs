// Title: Insert a PNG picture into cell B2 and rotate it 90° clockwise with Aspose.Cells for .NET
// AI Prompts: Add an image from a file path to cell B2, set its RotationAngle to 90 degrees, and save the workbook using Aspose.Cells in C#. | Place a PNG picture on a worksheet, apply a 90‑degree clockwise rotation, and export the file as an XLSX with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells how to add an image to a specific cell and rotate it | set rotation angle for picture inserted in Excel using Aspose.Cells .NET | insert PNG into Excel worksheet cell B2 with Aspose.Cells and rotate clockwise
// Tags: add picture to worksheet cell Aspose.Cells C# | apply 90-degree rotation to Excel picture Aspose.Cells | insert PNG image into Excel using Aspose.Cells | picture rotation property Aspose.Cells .NET | place image in cell B2 Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a PNG file into cell B2, rotates the picture 90 degrees clockwise, and saves the workbook as Output.xlsx.
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

            // Path to the image file to be inserted
            string imagePath = @"C:\Images\sample.png";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to cell B2 (row index 1, column index 1)
            int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);

            // Retrieve the picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Rotate the picture 90 degrees clockwise
            picture.RotationAngle = 90;

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
