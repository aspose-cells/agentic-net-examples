// Title: Resize an embedded PNG picture in an Excel worksheet to 800 × 600 pixels using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, retrieve the first picture on a worksheet, and assign 800 to Picture.Width and 600 to Picture.Height using Aspose.Cells in C#. | Programmatically scale an embedded PNG image in an Excel sheet to 800 × 600 pixels with the Aspose.Cells API. | Change the dimensions of a worksheet picture by setting pixel values on its Width and Height properties in a .NET project.
// Common Searches: how to change size of an embedded picture in Excel using Aspose.Cells C# | Aspose.Cells set picture dimensions in pixels for PNG image | resize worksheet image to 800x600 with Aspose.Cells .NET | C# code to adjust Excel picture width and height using Aspose.Cells
// Tags: Aspose.Cells picture resize PNG | set picture dimensions Aspose.Cells C# | worksheet image scaling .NET | Picture.Width Height property Aspose.Cells | embedded PNG size adjustment Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, accesses the first worksheet, checks for a picture, sets its Width to 800 and Height to 600 pixels via the Picture.Width and Picture.Height properties, and saves the workbook with the resized image.
class ResizeWorksheetImage
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook that contains the PNG image
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one picture on the worksheet
            if (sheet.Pictures.Count > 0)
            {
                // Get the first picture (the PNG image)
                Picture picture = sheet.Pictures[0];

                // Set the desired size in pixels (Aspose.Cells uses points; 1 point ≈ 1 pixel for screen display)
                picture.Width = 800;   // width in pixels
                picture.Height = 600;  // height in pixels
            }
            else
            {
                Console.WriteLine("No pictures found on the worksheet.");
            }

            // Save the workbook with the resized image
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
