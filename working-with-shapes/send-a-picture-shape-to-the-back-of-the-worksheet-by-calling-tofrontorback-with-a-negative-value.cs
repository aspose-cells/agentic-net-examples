// Title: Place a picture shape behind all other objects in an Excel worksheet using Aspose.Cells for .NET (ToFrontOrBack with negative value)
// Common Searches: how to send a picture shape to the back of an Excel worksheet using Aspose.Cells C# | Aspose.Cells picture ToFrontOrBack(-1) example | move image shape behind other objects in .NET Excel file with Aspose.Cells | set picture shape z-order to back in Aspose.Cells workbook | C# code to add a PNG picture and send it to back in Excel using Aspose.Cells
// Tags: picture shape layering Aspose.Cells | ToFrontOrBack negative argument .NET | send picture to back Excel worksheet | Aspose.Cells shape z-order control | add picture shape and reorder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, checks for a PNG file, inserts it as a picture shape at cell B2, calls picture.ToFrontOrBack(-1) to move the image behind all other worksheet objects, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            const string imagePath = "example.png";

            // Verify that the image file exists before adding it
            if (File.Exists(imagePath))
            {
                // Add a picture shape to the worksheet
                int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Send the picture to the back of the worksheet by using a negative value
                picture.ToFrontOrBack(-1);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
