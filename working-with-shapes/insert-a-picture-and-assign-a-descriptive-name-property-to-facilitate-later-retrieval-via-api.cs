// Title: C# – Insert a picture into Excel and assign a custom Name with Aspose.Cells
// Description: Demonstrates how to create a workbook, add an image to cell B2, give the picture a descriptive Name (e.g., "CompanyLogo"), set alternative text for accessibility, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells add picture C# | set picture name Aspose.Cells | Excel image insertion .NET | picture alternative text Aspose | retrieve picture by name Aspose.Cells
// Common Searches: Aspose.Cells insert image into worksheet | How to name a picture in Aspose.Cells | Find picture by Name property in Excel using Aspose | Set alt text for Excel picture with Aspose.Cells
// Developer Intent: Add an image to a worksheet and give it a unique Name for later API‑based retrieval or modification.
// Use Cases: Place a company logo at B2, name it "CompanyLogo", and later replace or resize it via the Name property. | Provide alt text for embedded graphics to comply with accessibility guidelines in generated reports. | Search for a picture by its assigned Name to programmatically adjust its position, size, or source image.
// AI Prompts: Write C# code that inserts a picture into cell B2, sets its Name to "HeaderImage", and saves the workbook with Aspose.Cells. | Show how to locate a picture by Name in an existing workbook and change its height and width using Aspose.Cells for .NET. | Explain a script to batch rename all pictures in a worksheet following a "Img_001", "Img_002" pattern with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureDemo
{
    // Demonstrates how to create a workbook, add an image to cell B2, give the picture a descriptive Name (e.g., "CompanyLogo"), set alternative text for accessibility, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "sample.jpg";

                // Verify that the image file exists before attempting to add it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Please ensure the file exists.");
                    return;
                }

                // Add the picture to the worksheet at cell position (row 2, column 2)
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);

                // Retrieve the added picture object
                Picture picture = worksheet.Pictures[pictureIndex];

                // Assign a descriptive name to the picture for later retrieval
                picture.Name = "CompanyLogo";

                // Optionally, set alternative text for accessibility
                picture.AlternativeText = "Company logo displayed in the report";

                // Save the workbook
                string outputPath = "PictureWithName.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
