// Title: How to unlock every shape on an Excel worksheet using Aspose.Cells for .NET by iterating the Shape collection
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, loops through worksheet.Shapes, sets each shape's IsLocked property to false, and saves the modified file. | Update an existing Aspose.Cells .NET project to remove the locked flag from all shapes on a specific worksheet without altering other worksheet content. | Create a console utility that accepts input and output Excel file paths, unlocks all shapes on the first sheet using Aspose.Cells, and gracefully handles missing file errors.
// Common Searches: Aspose.Cells C# unlock all shapes in a worksheet | remove shape protection from Excel file using Aspose.Cells .NET | iterate through ShapeCollection to set IsLocked false Aspose.Cells | batch unlock worksheet shapes programmatically in C# | how to disable shape locking in Aspose.Cells workbook
// Tags: unlock worksheet shapes Aspose.Cells | iterate ShapeCollection C# | set IsLocked false Excel shape | Aspose.Cells shape protection removal | batch unlock shapes workbook

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, iterates over its Shapes collection, clears the IsLocked flag for each shape, and saves the result to a new file while handling missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unlock all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Correct property to unlock a shape
                shape.IsLocked = false;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
