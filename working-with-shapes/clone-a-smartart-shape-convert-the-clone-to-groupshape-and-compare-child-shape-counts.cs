// Title: Clone a SmartArt shape in an Excel worksheet with Aspose.Cells for .NET and compare shape counts
// AI Prompts: Write C# code that loads an Excel file, finds the first SmartArt shape on the first worksheet, clones it with AddCopy, and prints the number of shapes before and after the clone using Aspose.Cells. | Show how to check the type of a cloned shape and read its dimensions after duplication with Aspose.Cells for .NET.
// Common Searches: how to duplicate a SmartArt shape in an Excel file using Aspose.Cells C# | Aspose.Cells AddCopy example for cloning worksheet shapes | count worksheet shapes before and after cloning with Aspose.Cells | retrieve shape collection size in Aspose.Cells after adding a copy
// Tags: clone smartart shape aspnet cells | addcopy shape duplication aspnet cells | worksheet shape count before after aspnet cells | smartart shape type verification aspnet cells | excel workbook shape cloning aspnet cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, selects the first shape on the first worksheet, clones it using the AddCopy method, prints the shape count before and after cloning, and saves the updated file.
class SmartArtCloneExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first shape on the worksheet
            Shape firstShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                firstShape = shape;
                break;
            }

            if (firstShape == null)
            {
                Console.WriteLine("No shape found on the worksheet.");
                return;
            }

            // Record shape count before cloning
            int countBefore = sheet.Shapes.Count;

            // Clone the shape using AddCopy with position and size parameters
            Shape clonedShape = sheet.Shapes.AddCopy(
                firstShape,
                firstShape.UpperLeftRow,
                firstShape.UpperLeftColumn,
                (int)firstShape.Height,
                (int)firstShape.Width);

            // Record shape count after cloning
            int countAfter = sheet.Shapes.Count;

            // Output the shape counts before and after cloning
            Console.WriteLine($"Original shape count: {countBefore}");
            Console.WriteLine($"Cloned shape count: {countAfter}");

            // Save the workbook with the cloned shape
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
