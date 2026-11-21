// Title: How to log each SmartArt shape’s type, row/column position, and size in an Excel file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get SmartArt shape type and dimensions from worksheet | retrieve row and column of each shape in Excel using Aspose.Cells | log shape width and height in points with Aspose.Cells C# | enumerate all shapes including SmartArt in a workbook using Aspose.Cells | Aspose.Cells example to print shape coordinates and size to console | how to iterate worksheet shapes and output their properties in C#
// Tags: Aspose.Cells enumerate worksheet shapes C# | retrieve SmartArt shape dimensions Aspose.Cells | log Excel shape type position size | extract shape coordinates Aspose.Cells API | SmartArt shape properties Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, loops through every worksheet, accesses the Shapes collection, and for each shape (including SmartArt) writes its type, upper‑left row and column, width, and height (in points) to the console, then optionally saves the file.
class SmartArtInfoLogger
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Loop through each shape
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Retrieve position information (row/column indices)
                    int upperLeftRow = shape.UpperLeftRow;
                    int upperLeftColumn = shape.UpperLeftColumn;

                    // Retrieve size information (in points)
                    double width = shape.Width;
                    double height = shape.Height;

                    // Log detailed information
                    Console.WriteLine($"Worksheet: {sheet.Name}");
                    Console.WriteLine($"Shape Index: {i}");
                    Console.WriteLine($"Type: {shape.Type}");
                    Console.WriteLine($"Position - Row: {upperLeftRow}, Column: {upperLeftColumn}");
                    Console.WriteLine($"Size - Width: {width} pts, Height: {height} pts");
                    Console.WriteLine(new string('-', 50));
                }
            }

            // Save the workbook (optional, if you made changes)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
