// Title: Read a shape's glow color from an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: how to retrieve the glow color of a shape in an Excel file using Aspose.Cells C# | Aspose.Cells C# read shape effectformat glow property | C# get visual effect color of a worksheet shape via Aspose.Cells | Aspose.Cells shape effectformat not accessible how to work around | extract shape glow settings from .xlsx using Aspose.Cells .NET
// Tags: read shape glow color Aspose.Cells | Aspose.Cells shape effectformat .NET | Excel worksheet shape visual effect extraction C# | C# retrieve shape properties Aspose.Cells | Aspose.Cells limitation effectformat access

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, accesses the first worksheet, and obtains a shape by name or index. It logs the shape's name and type, and explains that Aspose.Cells for .NET does not expose the EffectFormat (glow) property, so the current glow color cannot be read directly.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape by name or index.
            // Example using name "MyShape". Change as needed.
            Shape shape = sheet.Shapes["MyShape"]; // or sheet.Shapes[0] for index

            if (shape != null)
            {
                // Aspose.Cells Shape does not expose EffectFormat (glow) directly.
                // Output basic shape information.
                Console.WriteLine($"Shape found: Name = '{shape.Name}', Type = {shape.Type}");
            }
            else
            {
                Console.WriteLine("Shape not found.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
