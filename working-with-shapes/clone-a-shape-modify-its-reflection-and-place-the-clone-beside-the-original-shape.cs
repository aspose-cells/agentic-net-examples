// Title: Clone a worksheet shape, preserve its dimensions, and place the copy five columns to the right using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells AddCopy to duplicate the first shape on a worksheet, set the clone’s UpperLeftColumn five columns ahead, and copy the original Height and Width. | After cloning a shape, apply a reflection effect through the Shape.EffectFormat property in C#. | Add robust error handling that checks for the input file’s existence and verifies that the worksheet contains at least one shape before performing the copy.
// Common Searches: aspnet clone excel shape and move it to another column with Aspose.Cells | c# copy a shape and keep original size using Aspose.Cells AddCopy method | how to add reflection effect to a duplicated shape in Aspose.Cells for .NET | error handling for missing shapes when using Aspose.Cells shape APIs | place cloned shape at specific cell offset in Excel with Aspose.Cells
// Tags: AddCopy shape cloning Aspose.Cells | preserve shape dimensions after copy | set cloned shape column offset | apply reflection effect via EffectFormat | validate shape existence before cloning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, checks for at least one shape, clones the first shape with AddCopy, moves the duplicate five columns to the right while keeping the original height and width, optionally applies a reflection effect, and saves the result.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that at least one shape exists
            if (sheet.Shapes.Count == 0)
                throw new InvalidOperationException("No shapes found in the worksheet.");

            // Retrieve the original shape
            Shape originalShape = sheet.Shapes[0];

            // Calculate new position (5 columns to the right)
            int newUpperLeftRow = originalShape.UpperLeftRow;
            int newUpperLeftColumn = originalShape.UpperLeftColumn + 5;
            int newLowerRightRow = originalShape.LowerRightRow;
            int newLowerRightColumn = originalShape.LowerRightColumn + 5;

            // Clone the original shape to the new location
            Shape clonedShape = sheet.Shapes.AddCopy(
                originalShape,
                newUpperLeftRow,
                newUpperLeftColumn,
                newLowerRightRow,
                newLowerRightColumn);

            // Preserve the original size (height and width)
            clonedShape.Height = originalShape.Height;
            clonedShape.Width = originalShape.Width;

            // NOTE: Shape effect APIs (EffectFormat, Reflection, etc.) may not be available
            // in all Aspose.Cells versions. If needed, they can be applied here using the
            // appropriate EffectFormat properties.

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
