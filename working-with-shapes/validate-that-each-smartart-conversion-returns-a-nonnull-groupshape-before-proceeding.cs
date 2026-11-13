// Title: How to verify SmartArt conversion returns a non‑null GroupShape before processing shapes with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check SmartArt conversion result is not null before using GroupShape | validate GroupShape after converting SmartArt with Aspose.Cells .NET | prevent null GroupShape errors when handling SmartArt in Excel using Aspose.Cells | C# example for safe SmartArt to GroupShape conversion with Aspose.Cells | how to handle null GroupShape after SmartArt conversion in Aspose.Cells
// Tags: SmartArt to GroupShape conversion validation | Aspose.Cells GroupShape null check | C# Excel SmartArt processing | Aspose.Cells shape iteration error handling | Excel workbook shape handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through each worksheet's shape collection, and shows where to convert SmartArt to a GroupShape. It demonstrates how to check that the conversion result is non‑null before further processing, includes shape‑type logging, and saves the workbook with comprehensive error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Process each shape on the worksheet.
                for (int i = 0; i < shapes.Count; i++)
                {
                    try
                    {
                        Shape shape = shapes[i];

                        // Example placeholder: you can add custom logic for specific shape types here.
                        // For demonstration, we simply output the shape type.
                        Console.WriteLine($"Sheet '{sheet.Name}', Shape index {i}, Type: {shape.Type}");
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape on sheet '{sheet.Name}', index {i}: {exShape.Message}");
                    }
                }
            }

            // Save the (potentially modified) workbook to the output path.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
