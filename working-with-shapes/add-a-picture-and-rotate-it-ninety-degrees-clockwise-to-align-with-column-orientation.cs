// Title: Insert a PNG image into an Excel worksheet at cell A1 and rotate it 90° clockwise using Aspose.Cells for .NET (C#)
// AI Prompts: Insert an image file into cell A1 of a new workbook and set its RotationAngle property to 90 degrees with Aspose.Cells in C#. | Add a picture to a worksheet, rotate it ninety degrees clockwise, and save the workbook using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to add a picture to a specific cell and rotate it | rotate inserted image 90 degrees using Aspose.Cells .NET | Aspose.Cells picture rotation example for Excel workbook | insert PNG into Excel sheet programmatically with Aspose.Cells C# | set picture rotation angle in Aspose.Cells workbook
// Tags: picture insertion at cell A1 Aspose.Cells | image rotation property Aspose.Cells C# | PNG picture embed Excel Aspose.Cells | worksheet picture manipulation Aspose.Cells | save workbook after picture rotation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a PNG image at cell A1, rotates the picture 90° clockwise via the RotationAngle property, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "image.png";

            // Insert picture only if the file exists
            if (File.Exists(imagePath))
            {
                // Add the picture to the worksheet at cell A1 (top‑left corner)
                // The Add method returns the index of the inserted picture
                int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Rotate the picture 90 degrees clockwise to align with column orientation
                picture.RotationAngle = 90;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("Output.xlsx");
            Console.WriteLine("Workbook saved successfully as Output.xlsx.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
