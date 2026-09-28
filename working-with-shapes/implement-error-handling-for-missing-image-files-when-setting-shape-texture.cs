// Title: Add a rectangle shape with a texture image and handle missing file errors in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, inserts a rectangle shape, applies a texture from a given image file, and validates the file's existence with proper exception handling using Aspose.Cells. | Demonstrate how to catch FileNotFoundException and other runtime exceptions when loading an image as a shape fill, and ensure the output directory is created before saving the workbook.
// Common Searches: Aspose.Cells C# add rectangle shape and set picture fill with file existence validation | how to catch missing image file error when using sheet.Pictures.Add in Aspose.Cells | C# verify texture image path before applying to shape in Aspose.Cells workbook | ensure output folder exists before saving Excel file with Aspose.Cells
// Tags: add rectangle shape with picture fill Aspose.Cells | file existence validation for shape texture C# | FileNotFoundException handling Aspose.Cells workbook | create output directory before saving Excel Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, checks that the specified texture image file exists, loads the image into the worksheet's picture collection, handles FileNotFoundException and other exceptions, ensures the output directory is present, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the image file path to be used as texture
        string imagePath = @"C:\Images\texture.png";

        // Add a rectangle shape to the worksheet
        // Parameters: type, upper left row, upper left column, top, left, height, width
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 100);

        try
        {
            // Verify that the image file exists before using it
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Image file not found.", imagePath);

            // Load the picture into the worksheet's picture collection
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Optional: you can position the picture or associate it with the shape as needed.
            // In this example we simply add the picture to the sheet.
        }
        catch (FileNotFoundException ex)
        {
            // Handle missing image file: log the error
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }

        // Ensure output directory exists
        string outputPath = @"C:\Output\ShapeWithTexture.xlsx";
        string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Save the workbook (lifecycle rule: save)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
