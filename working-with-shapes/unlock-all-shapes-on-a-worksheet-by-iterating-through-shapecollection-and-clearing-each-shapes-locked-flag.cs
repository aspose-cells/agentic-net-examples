// Title: Unlock every shape in an Excel worksheet using Aspose.Cells for .NET by iterating the ShapeCollection
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, loops through all shapes on a worksheet, sets each Shape.IsLocked to false, and saves the updated file. | Show a step‑by‑step example of removing the locked flag from every shape on a specific worksheet using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# unlock all shapes in a worksheet | remove shape protection from an Excel file using Aspose.Cells .NET | iterate through ShapeCollection to set IsLocked false in Aspose.Cells | how to programmatically unlock shapes in an existing Excel workbook with C# | Aspose.Cells example for clearing shape locked property
// Tags: unlock worksheet shapes Aspose.Cells | iterate ShapeCollection C# | clear Shape.IsLocked property | modify Excel shapes programmatically | save workbook after shape changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, iterates over every shape in the first worksheet, sets each shape's IsLocked property to false, and saves the modified workbook to a new file.
class UnlockShapesExample
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired worksheet name/index)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet and unlock them
            foreach (Shape shape in sheet.Shapes)
            {
                // Use the correct property to unlock the shape
                shape.IsLocked = false;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
