// Title: C# – Insert a Picture into an Aspose.Cells Worksheet and Assign a Custom Name
// Description: Creates a new workbook, adds an image from a file to cell B2, sets the picture's Name property (e.g., "CompanyLogo"), includes file‑existence validation, and saves the file as an XLSX document.
// Keywords: Aspose.Cells add image C# | set picture name Aspose.Cells | retrieve picture by name .NET | Excel picture insertion Aspose | C# workbook image handling | global
// Common Searches: how to add an image to an Aspose.Cells worksheet with a name | Aspose.Cells C# picture Name property example | retrieve a shape by Name in Aspose.Cells | error handling for missing picture file Aspose.Cells
// Developer Intent: Add an image to a worksheet and give it a unique Name so it can be identified later via the Aspose.Cells API.
// Use Cases: Place a company logo at B2, name it "CompanyLogo", and later replace or hide it programmatically. | Insert a watermark, assign a distinct name, and toggle its visibility based on user settings. | Add a diagram, set a descriptive name, and update the diagram image in future runs without searching by index.
// AI Prompts: Write C# code using Aspose.Cells to insert a picture from a file path, set its Name property, and handle missing‑file errors. | Show how to locate a picture by its Name in an existing workbook and modify its position or replace the image with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds an image from a file to cell B2, sets the picture's Name property (e.g., "CompanyLogo"), includes file‑existence validation, and saves the file as an XLSX document.
class InsertPictureWithName
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        const string imagePath = "sample.jpg";

        try
        {
            // Verify that the image file exists before attempting to add it
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file '{imagePath}' was not found.");

            // Add a picture to the worksheet (top-left corner at row 2, column 2)
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Assign a descriptive name to the picture for later retrieval
            picture.Name = "CompanyLogo";
        }
        catch (Exception ex)
        {
            // Log the error; the workbook will be saved without the picture
            Console.WriteLine($"Warning: Unable to add picture. {ex.Message}");
        }

        // Save the workbook
        try
        {
            workbook.Save("PictureWithName.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
