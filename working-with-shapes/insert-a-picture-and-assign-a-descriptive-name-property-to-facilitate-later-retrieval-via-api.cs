// Title: How to insert a picture into an Excel worksheet and assign a custom Name property using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert image into Excel and set picture.Name for later lookup | C# code to add a PNG to cell A1 and give it a descriptive name with Aspose.Cells | Retrieve a specific picture from an Aspose.Cells workbook by custom Name property | Set custom identifier on inserted picture in Aspose.Cells .NET example | Avoid FileNotFoundException when adding a picture with Aspose.Cells C#
// Tags: add picture to Excel worksheet Aspose.Cells C# | assign custom Name to picture Aspose.Cells | picture retrieval by Name property Aspose.Cells | insert PNG at cell A1 using Aspose.Cells | Aspose.Cells image handling C# example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample checks for a PNG file, creates a new workbook, inserts the image into cell A1, assigns the picture a descriptive Name ("CompanyLogo") for later API retrieval, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Path to the image file to be inserted
        string imagePath = "sample.png";

        // Verify that the image file exists to avoid FileNotFoundException
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Error: Image file not found at path '{imagePath}'.");
            return;
        }

        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert the picture at cell A1 (row 0, column 0)
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Retrieve the inserted picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Assign a descriptive name to the picture for later retrieval via API
            picture.Name = "CompanyLogo";

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
