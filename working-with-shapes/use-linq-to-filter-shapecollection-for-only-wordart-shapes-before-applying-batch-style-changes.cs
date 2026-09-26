// Title: Use LINQ to select only WordArt shapes from an Excel worksheet and apply fill and font styling with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, uses LINQ to retrieve shapes where IsWordArt is true, and sets each shape's FillFormat.ForeColor to LightBlue and Font to bold Arial 14. | Generate a method that accepts a Worksheet object, filters its ShapeCollection for WordArt shapes using LINQ, and applies a light blue fill and Arial 14 bold font to the shapes.
// Common Searches: aspnet linq filter wordart shapes in worksheet using aspose.cells | how to change fill color of WordArt objects in Excel with Aspose.Cells C# | batch update font properties of WordArt shapes in Aspose.Cells | select only WordArt shapes from ShapeCollection with LINQ in C# | apply style changes to multiple WordArt shapes in an Excel file programmatically
// Tags: linq filter wordart shapes aspose.cells | batch style update wordart excel c# | shapecollection wordart selection aspose | fillformat color change wordart c# | font formatting wordart aspose.cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

// Loads an Excel workbook, uses LINQ to extract only WordArt shapes from the first worksheet's ShapeCollection, changes each shape's fill color to LightBlue and sets the font to bold Arial 14, then saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Select only WordArt shapes from the worksheet's ShapeCollection
            List<Shape> wordArtShapes = sheet.Shapes
                                            .Cast<Shape>()
                                            .Where(s => s.IsWordArt)
                                            .ToList();

            // Apply batch style changes to each WordArt shape
            foreach (Shape wordArt in wordArtShapes)
            {
                // Change fill color using FillFormat
                wordArt.FillFormat.ForeColor = Color.LightBlue;

                // Modify font properties
                wordArt.Font.Name = "Arial";
                wordArt.Font.Size = 14;
                wordArt.Font.IsBold = true;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
