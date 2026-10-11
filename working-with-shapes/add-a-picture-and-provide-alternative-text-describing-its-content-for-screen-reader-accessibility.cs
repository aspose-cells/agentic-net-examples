// Title: Insert a JPEG picture into cell A1 of an Excel workbook and assign alternative text for screen‑reader accessibility with Aspose.Cells for .NET
// AI Prompts: Insert an image from a local file into cell A1 of a new workbook and assign a descriptive AlternativeText property using Aspose.Cells in C#. | Update the code to place a picture in a custom cell range and assign its alternative text from a variable or resource file.
// Common Searches: Aspose.Cells C# how to embed a picture in an Excel sheet and add alt description for screen readers | provide accessibility text for image inserted with Aspose.Cells .NET | add image to specific cell range in Excel using Aspose.Cells and ensure alt text is set | save Excel file with embedded JPEG and accessibility metadata using Aspose.Cells for .NET | check image file existence before adding picture with Aspose.Cells C#
// Tags: insert picture into worksheet Aspose.Cells C# | assign alt description to Excel image Aspose.Cells | embed JPEG in Excel workbook Aspose.Cells | accessibility for Excel pictures Aspose.Cells | save workbook with images Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a new workbook, embed a JPEG picture into cell A1, set a meaningful AlternativeText for screen‑reader accessibility, and save the file as an .xlsx using Aspose.Cells for .NET.
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

            // Path to the image file
            string imagePath = @"C:\Images\sample.jpg";

            // Verify the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add a picture to cell A1 (row 0, column 0)
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Retrieve the picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Provide alternative text for screen reader accessibility
            picture.AlternativeText = "A sample photograph showing a sunrise over mountains.";

            // Ensure the output directory exists
            string outputPath = "OutputWithPicture.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
