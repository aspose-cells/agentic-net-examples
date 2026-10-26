// Title: Log SmartArt shape type, position, and size from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get SmartArt shape coordinates and dimensions | list SmartArt objects in an Excel file with Aspose.Cells .NET | retrieve SmartArt type and size from a worksheet using Aspose.Cells | Aspose.Cells enumerate shapes and filter SmartArt in C# | how to output SmartArt shape details from an Excel workbook in .NET
// Tags: Aspose.Cells enumerate SmartArt shapes | SmartArt shape metadata extraction .NET | retrieve SmartArt dimensions Excel | log SmartArt type and position C# | shape.IsSmartArt property usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through each worksheet and its shapes, identifies SmartArt shapes via the IsSmartArt property, and writes to the console the worksheet name, SmartArt name, type, upper‑left row and column, and height and width in points.
class SmartArtInfoLogger
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Use the IsSmartArt property to identify SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Log detailed information about the SmartArt shape
                        Console.WriteLine($"Worksheet: {sheet.Name}");
                        Console.WriteLine($"SmartArt Name: {shape.Name}");
                        Console.WriteLine($"SmartArt Type: {shape.Type}");
                        Console.WriteLine($"Position - Upper Left Row: {shape.UpperLeftRow}, Upper Left Column: {shape.UpperLeftColumn}");
                        Console.WriteLine($"Size - Height: {shape.Height} points, Width: {shape.Width} points");
                        Console.WriteLine(new string('-', 50));
                    }
                }
            }

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
