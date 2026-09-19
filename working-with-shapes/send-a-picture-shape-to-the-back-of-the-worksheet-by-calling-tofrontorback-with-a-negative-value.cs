// Title: Send a picture shape to the back of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert an image into a worksheet and programmatically send it behind all other objects with Aspose.Cells in C#. | Call Picture.ToFrontOrBack(-1) to move a picture shape to the back layer of an Excel sheet.
// Common Searches: Aspose.Cells C# move picture shape to back layer | how to use ToFrontOrBack with negative value in Aspose.Cells | place image behind cells using Aspose.Cells for .NET | send picture shape behind other shapes in Excel with Aspose.Cells | C# Aspose.Cells picture ordering back front example
// Tags: Aspose.Cells picture ToFrontOrBack method | C# picture back placement | Aspose.Cells embed picture behind worksheet cells | Excel picture shape layering Aspose | Aspose.Cells picture ordering negative value

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example creates a new workbook, adds a picture from 'image.png' at cell (5,5), calls picture.ToFrontOrBack(-1) to send the picture to the back of the sheet, and saves the workbook as 'Result.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                string imagePath = "image.png";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add a picture shape to the worksheet
                    int pictureIndex = worksheet.Pictures.Add(5, 5, imagePath);

                    // Retrieve the added picture shape
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Send the picture to the back of the worksheet
                    picture.ToFrontOrBack(-1);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // Save the workbook
                string resultPath = "Result.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
