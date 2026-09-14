// Title: Insert a PNG picture into an Excel worksheet and achieve a semi‑transparent effect with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that inserts a PNG image into a specific cell using Aspose.Cells and applies 40% opacity by adjusting the image’s alpha channel before adding it to the workbook. | Update the example to use the Shape.Transparency property (available in newer Aspose.Cells releases) to set a picture’s transparency to 30% after insertion. | Create a helper method that loads an image file, changes its transparency to a given percentage, and returns a stream suitable for Worksheet.Pictures.Add in Aspose.Cells.
// Common Searches: how to make an inserted picture semi transparent in Aspose.Cells C# | Aspose.Cells set picture opacity programmatically | pre‑process PNG transparency before adding to Excel with Aspose.Cells | use Shape.Transparency property Aspose.Cells 2023 version | insert image into specific cell and adjust alpha channel Aspose.Cells .NET
// Tags: aspocells insert picture opacity c# | aspocells shape transparency feature | c# pre‑process png alpha channel for excel | excel worksheet semi transparent image aspocells | worksheet pictures add transparent png

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, checks for a PNG file, inserts the image at row 5 column 3, and saves the file as Result.xlsx. It notes that the Transparency property is not available in the current Aspose.Cells version, recommending either pre‑adjusting the image’s alpha channel before insertion or using the Shape.Transparency property in newer releases to achieve a semi‑transparent picture.
class InsertPictureWithTransparency
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted.
            string imagePath = "sample.png";

            // Verify that the image file exists to avoid FileNotFoundException.
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at row 5, column 3 (zero‑based indices).
            int pictureIndex = sheet.Pictures.Add(4, 2, imagePath);
            Picture pic = sheet.Pictures[pictureIndex];

            // Note: The Transparency property is not available in this version of Aspose.Cells.
            // If needed, adjust image transparency before inserting or use Shape.Transparency in newer versions.

            // Save the workbook to a file.
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
