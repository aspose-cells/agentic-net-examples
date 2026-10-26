// Title: Locate the worksheet that contains a named shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find which worksheet a shape belongs to by shape name | search for a specific shape across all sheets in an Excel workbook with Aspose.Cells | determine worksheet containing a shape named MyShape using Aspose.Cells .NET | iterate worksheets to locate a shape by name in Aspose.Cells C# example | how to get worksheet reference for a shape in an Excel file using Aspose.Cells
// Tags: iterate worksheets Aspose.Cells | shape name lookup Aspose.Cells | retrieve worksheet for shape Aspose.Cells | search shapes across worksheets C# | find shape location Excel Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads an Excel workbook, defines the target shape name, loops through each worksheet and its Shapes collection, compares each shape's Name property to the target, captures the worksheet where the match is found, and prints the worksheet name or a not‑found message.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Define the name of the shape to locate
        string targetShapeName = "MyShape";

        // Variable to hold the worksheet that contains the target shape
        Worksheet targetWorksheet = null;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the current worksheet
            foreach (Aspose.Cells.Drawing.Shape shape in sheet.Shapes)
            {
                // Check if the shape's name matches the target name
                if (shape.Name == targetShapeName)
                {
                    targetWorksheet = sheet;
                    break; // Shape found, exit inner loop
                }
            }

            if (targetWorksheet != null)
                break; // Shape found, exit outer loop
        }

        // Output the result
        if (targetWorksheet != null)
        {
            Console.WriteLine($"Shape '{targetShapeName}' found in worksheet: {targetWorksheet.Name}");
        }
        else
        {
            Console.WriteLine($"Shape '{targetShapeName}' not found in any worksheet.");
        }
    }
}
