// Title: Count Worksheets Containing Shapes with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, walks through each worksheet, checks the Shapes collection, and prints the number of sheets that have at least one drawing object. Demonstrates quick reporting of shape presence using Aspose.Cells.
// Keywords: Aspose.Cells | C# | worksheet shape count | Excel shapes detection | count sheets with drawings | Aspose.Cells Shapes API | enumerate worksheet graphics | Excel workbook analysis
// Common Searches: how to count worksheets with shapes using Aspose.Cells | C# code to find Excel sheets that contain drawings | Aspose.Cells get number of sheets that have graphics | report worksheets containing shapes in a workbook | enumerate shape objects across Excel worksheets C#
// Developer Intent: Identify how many worksheets in a workbook contain at least one shape.
// Use Cases: Generate a summary that shows how many sheets include charts, images, or other drawing objects. | Validate that required graphics are present before publishing or converting a workbook. | Trigger downstream processing only for worksheets that contain shapes, such as exporting them separately.
// AI Prompts: Create a method that returns the names of all worksheets with shapes using Aspose.Cells for .NET. | Extend the sample to also total the number of shapes across the workbook and display both counts. | Write a console app that logs worksheet names containing shapes and copies those sheets into a new workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsShapeCountDemo
{
    // Loads an Excel workbook, walks through each worksheet, checks the Shapes collection, and prints the number of sheets that have at least one drawing object. Demonstrates quick reporting of shape presence using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the workbook file path (use first argument or default)
            string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

            // Load the workbook (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(workbookPath);

            int worksheetsWithShapes = 0;

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet contains any shapes
                if (sheet.Shapes.Count > 0)
                {
                    worksheetsWithShapes++;
                }
            }

            // Output the count of worksheets that contain shapes
            Console.WriteLine($"Number of worksheets containing shapes: {worksheetsWithShapes}");

            // Optionally, save the workbook (demonstrates the provided Save rule)
            // workbook.Save("output.xlsx");
        }
    }
}
