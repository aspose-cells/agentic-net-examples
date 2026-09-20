// Title: How to duplicate a picture shape and assign a new texture image in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that clones an existing picture shape in a worksheet, loads a different PNG file as its image source, moves the clone to a non‑overlapping position, and saves the workbook with Aspose.Cells. | Write a method that copies a shape's size and location, adds a new picture from a file stream, replaces the original texture, and updates the Excel file using the Aspose.Cells API.
// Common Searches: asp.net duplicate picture shape in Excel workbook with Aspose.Cells | c# replace image of copied shape in Aspose.Cells worksheet | how to add a new texture to a cloned shape using Aspose.Cells for .NET | Aspose.Cells example: copy shape and change its PNG source | programmatically reposition a duplicated picture shape in Excel C#
// Tags: duplicate picture shape Aspose.Cells | set shape image from file stream C# | reposition cloned shape worksheet | replace shape texture Aspose.Cells | add picture shape with custom PNG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, ensures a picture shape exists, clones that shape preserving its size and position, assigns a different PNG file as the texture for the copy, shifts the duplicated shape horizontally to avoid overlap, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string originalImgPath = "OriginalTexture.png";
            string newImgPath = "NewTexture.png";
            string outputPath = "Output.xlsx";

            // Verify required files exist
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Workbook not found: {inputPath}");
            if (!File.Exists(originalImgPath))
                throw new FileNotFoundException($"Original image not found: {originalImgPath}");
            if (!File.Exists(newImgPath))
                throw new FileNotFoundException($"New image not found: {newImgPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one picture shape
            if (sheet.Shapes.Count == 0)
            {
                using (FileStream fs = File.OpenRead(originalImgPath))
                {
                    // Add a picture shape using the original image
                    sheet.Shapes.AddPicture(1, 1, 100, 100, fs);
                }
            }

            // Get the first shape (assumed to be a picture)
            Shape originalShape = sheet.Shapes[0] as Shape;
            if (originalShape == null)
                throw new InvalidOperationException("The first shape is not a picture.");

            // Duplicate the shape by adding a new picture with the same size/position
            Shape duplicatedShape;
            using (FileStream fsNew = File.OpenRead(newImgPath))
            {
                duplicatedShape = sheet.Shapes.AddPicture(
                    originalShape.UpperLeftRow,
                    originalShape.UpperLeftColumn,
                    (int)originalShape.Height,
                    (int)originalShape.Width,
                    fsNew);
            }

            // Reposition the duplicated shape so it does not overlap the original
            duplicatedShape.Left = originalShape.Left + 150; // move 150 points to the right
            duplicatedShape.Top = originalShape.Top;         // keep same vertical position

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
