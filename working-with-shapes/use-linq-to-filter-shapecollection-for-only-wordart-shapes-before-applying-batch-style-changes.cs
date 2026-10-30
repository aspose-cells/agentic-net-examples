// Title: Filter WordArt shapes with LINQ and apply batch formatting using Aspose.Cells for .NET
// Common Searches: c# linq query to select only WordArt shapes from worksheet shape collection with Aspose.Cells | how to change fill color and line weight of multiple WordArt objects in an Excel file using Aspose.Cells | batch update text and style of WordArt shapes in a workbook with Aspose.Cells .NET | apply yellow fill and black border to all WordArt shapes in a spreadsheet using C# | filter shapes by type WordArt before formatting in Aspose.Cells
// Tags: linq filter shapecollection aspnet cells | batch format wordart shapes aspnet cells | set wordart fill color aspnet cells | modify wordart line weight aspnet cells | update wordart text aspnet cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

// Loads an Excel workbook, uses LINQ to retrieve only shapes whose Type contains "WordArt" from the first worksheet's ShapeCollection, then sets each shape's fill color to yellow, line weight to 2.0 with a black border, updates the displayed text, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Select only WordArt shapes from the worksheet's ShapeCollection
            var wordArtShapes = sheet.Shapes
                                    .Cast<Shape>()
                                    .Where(s => s.Type.ToString().Contains("WordArt"))
                                    .ToList();

            // Apply batch style changes to each WordArt shape
            foreach (var shape in wordArtShapes)
            {
                // Set fill color
                shape.FillFormat.ForeColor = Color.Yellow;

                // Set line weight and color
                shape.LineFormat.Weight = 2.0;
                shape.LineFormat.ForeColor = Color.Black;

                // Update displayed text
                shape.Text = "Updated WordArt";
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
