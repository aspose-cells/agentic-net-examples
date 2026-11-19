// Title: How to read a shape's glow color (ARGB) from an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read ARGB values of a shape's glow effect in Excel | Get glow color of a named shape in an .xlsx file with Aspose.Cells | Retrieve shape glow property using Aspose.Cells .NET API | Aspose.Cells example to read shape effect color in C# | Read shape glow attribute from workbook with Aspose.Cells
// Tags: Aspose.Cells shape glow extraction | C# read shape effect color from workbook | Excel shape glow property access | convert CellsColor to System.Drawing.Color | retrieve ARGB components of shape glow

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing .xlsx file, accesses the first worksheet, obtains a shape by its name, checks for a defined Glow effect, converts the Aspose.Cells CellsColor to a System.Drawing.Color, and writes the ARGB components of the glow color to the console, with handling for missing files or shapes.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape by name (replace "MyShape" with the actual shape name)
            Shape shape = sheet.Shapes["MyShape"];

            // Check that the shape exists and has a Glow effect defined
            if (shape != null && shape.Glow != null && shape.Glow.Color != null)
            {
                // Convert Aspose.Cells.CellsColor to System.Drawing.Color
                Color glowColor = shape.Glow.Color.Color;

                // Output ARGB components
                Console.WriteLine($"Glow Color - A:{glowColor.A}, R:{glowColor.R}, G:{glowColor.G}, B:{glowColor.B}");
            }
            else
            {
                Console.WriteLine("Shape not found or does not have a glow effect.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
