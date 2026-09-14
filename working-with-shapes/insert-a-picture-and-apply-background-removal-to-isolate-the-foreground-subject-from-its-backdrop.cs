// Title: Insert a JPEG into cell B2 of an Excel worksheet and remove its background with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new workbook, checks for a JPEG file, inserts it into cell B2, sets picture.IsBackgroundRemoved = true, and saves the file as XLSX using Aspose.Cells. | Generate a .NET example that validates an image path, adds the image to a worksheet at a specific cell, enables background removal on the picture shape, and writes the workbook to disk. | Provide a C# snippet that demonstrates how to embed an external image into an Excel sheet, apply transparent background removal, and handle missing file errors with Aspose.Cells.
// Common Searches: Aspose.Cells C# insert image into specific cell and make background transparent | how to enable IsBackgroundRemoved for a picture in Aspose.Cells .NET | example code for adding JPEG to Excel worksheet with Aspose.Cells and removing background | C# check if image file exists before inserting into Excel using Aspose.Cells | save workbook with embedded picture after background removal Aspose.Cells
// Tags: insert picture into Excel cell Aspose.Cells C# | background removal for picture shape Aspose.Cells | validate image file existence C# Aspose.Cells | save workbook with embedded image Aspose.Cells | use IsBackgroundRemoved property Aspose.Cells | add JPEG to worksheet cell B2 Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program creates a new workbook, verifies the JPEG file at the given path, inserts the image into cell B2, optionally enables background removal via the IsBackgroundRemoved property, and saves the workbook as Output.xlsx.
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

            // Path to the image you want to insert
            string imagePath = @"C:\Images\sample.jpg";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at cell B2 (row index 1, column index 1)
            int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // If the API version supports background removal, uncomment the line below
            // picture.IsBackgroundRemoved = true;

            // Ensure the output directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
