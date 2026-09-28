// Title: Duplicate Excel form control shapes from one worksheet to another while keeping linked cell addresses unchanged using Aspose.Cells for .NET
// AI Prompts: Copy all form control shapes from the source worksheet to a target worksheet in a workbook, preserving their linked cell references, using Aspose.Cells C#. | Create a new worksheet when it does not exist and duplicate each shape with the AddCopy method, keeping the original row and column bounds. | Catch shape‑copy exceptions, log any failures, and save the updated workbook to a separate file with Aspose.Cells.
// Common Searches: Aspose.Cells copy form controls to another sheet without changing linked cells | C# duplicate worksheet shapes preserving cell links using Aspose.Cells | How to use AddCopy to move Excel form controls between worksheets in .NET | Copy form control shapes from first worksheet to second worksheet Aspose.Cells example | Preserve linked cell addresses when copying Excel controls with Aspose.Cells
// Tags: AddCopy shape method Aspose.Cells .NET | duplicate form control shapes worksheet Aspose.Cells | preserve linked cell references Excel controls | copy shapes between worksheets C# | create target worksheet if missing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, selects the first worksheet as the source, ensures a target worksheet exists (or creates one), then iterates over each shape in the source sheet and copies it to the target sheet using the AddCopy method while retaining the original row/column positions and linked cell addresses. Any copy errors are caught and reported, and the modified workbook is saved to a new file.
class DuplicateFormControls
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

            // Source worksheet (first worksheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Target worksheet (second worksheet or create a new one)
            Worksheet targetSheet = workbook.Worksheets.Count > 1
                ? workbook.Worksheets[1]
                : workbook.Worksheets.Add("TargetSheet");

            // Iterate through all shapes in the source worksheet
            foreach (Shape shape in sourceSheet.Shapes)
            {
                try
                {
                    // Copy the shape to the target worksheet preserving its position
                    targetSheet.Shapes.AddCopy(
                        shape,
                        shape.UpperLeftRow,
                        shape.UpperLeftColumn,
                        shape.LowerRightRow,
                        shape.LowerRightColumn);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy shape '{shape.Name}': {ex.Message}");
                }
            }

            // Save the workbook with duplicated controls
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
