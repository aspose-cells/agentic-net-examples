// Title: Duplicate a shape, rename the copy, and offset it by 10 points in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that uses Aspose.Cells to locate a shape named "OriginalShape" on a worksheet, duplicate it with AddCopy, rename the duplicate to "OriginalShape_Copy", and shift its left position by 10 points. | Write a reusable C# method for Aspose.Cells that accepts a source shape name, a new name, and a horizontal offset, then copies the shape, assigns the new name, and moves it right by the specified points.
// Common Searches: Aspose.Cells C# duplicate shape and change its name on the same worksheet | C# offset copied Excel shape by specific points using Aspose.Cells | How to use AddCopy to clone a shape and move it in Aspose.Cells | Rename and reposition a shape copy in an Aspose.Cells workbook with C#
// Tags: addcopy shape Aspose.Cells C# | rename duplicated shape worksheet | horizontal offset shape Aspose.Cells | shape cloning Excel Aspose.Cells | shape positioning points Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, finds the shape "OriginalShape", creates a copy with AddCopy, renames the copy to "OriginalShape_Copy", moves it 10 points to the right, and saves the result as output.xlsx.
class ShapeDuplicationExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape named "OriginalShape"
            Shape originalShape = sheet.Shapes["OriginalShape"];
            if (originalShape == null)
            {
                Console.WriteLine("Shape 'OriginalShape' not found.");
                return;
            }

            // Duplicate the shape and add the copy to the same worksheet
            Shape copiedShape = sheet.Shapes.AddCopy(
                originalShape,
                originalShape.UpperLeftRow,
                originalShape.UpperLeftColumn,
                originalShape.LowerRightRow,
                originalShape.LowerRightColumn);

            copiedShape.Name = "OriginalShape_Copy";

            // Offset the copied shape 10 points to the right
            copiedShape.Left = originalShape.Left + 10;

            // Save the workbook with the duplicated shape
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
