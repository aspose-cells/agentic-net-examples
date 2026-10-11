// Title: Validate non‑null SmartArt GroupShape objects before conversion with Aspose.Cells for .NET
// AI Prompts: Loop through each worksheet, identify GroupShape objects, and throw an exception if any GroupShape instance is null. | Insert a null‑check for SmartArt GroupShape shapes before executing conversion logic in an Aspose.Cells C# application. | Create a safety routine that verifies SmartArt (GroupShape) objects are not null while processing the Shapes collection of every worksheet.
// Common Searches: C# Aspose.Cells how to ensure SmartArt GroupShape is not null before conversion | check for null SmartArt shapes in Excel workbook using Aspose.Cells .NET | validate GroupShape objects in each worksheet with Aspose.Cells to avoid NullReferenceException | Aspose.Cells iterate shapes and perform null safety for SmartArt in C#
// Tags: Aspose.Cells SmartArt GroupShape validation | C# null check for GroupShape before conversion | Excel workbook shape iteration Aspose.Cells | Prevent NullReferenceException SmartArt Aspose.Cells | GroupShape safety check .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, confirms its presence, and iterates through every worksheet's shape collection. For each shape that is a GroupShape (the Aspose.Cells representation of SmartArt), it verifies the instance is not null and throws an InvalidOperationException if a null reference is found. After validation, the workbook is saved to a new file.
class SmartArtConversionValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Iterate backwards when removing items to avoid index issues
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    Shape shape = shapes[i];

                    // In Aspose.Cells, SmartArt objects are represented as GroupShape
                    if (shape is GroupShape)
                    {
                        // Validation placeholder – ensure the shape instance is not null
                        if (shape == null)
                        {
                            throw new InvalidOperationException(
                                $"SmartArt shape is null on worksheet '{sheet.Name}', index {i}.");
                        }

                        // If a conversion were required, it would be performed here.
                        // Aspose.Cells does not provide a direct ConvertToGroupShape method,
                        // so we keep the original shape unchanged.
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
