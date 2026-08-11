// Title: Count Worksheets Containing Shapes with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook using Aspose.Cells, iterates through each worksheet, checks the Shapes collection, increments a counter for sheets with one or more shapes, and outputs the total number of worksheets that contain shapes.
// Keywords: Aspose.Cells | C# | count worksheets with shapes | Shapes collection | worksheet shape detection | Excel workbook analysis | Aspose.Cells API
// Common Searches: how to count sheets with shapes Aspose.Cells | C# get number of worksheets that have drawings | Aspose.Cells count worksheets containing images | retrieve worksheets that contain shapes .NET | Excel workbook shape count using Aspose
// Developer Intent: Determine how many worksheets in a workbook contain at least one shape.
// Use Cases: Create a summary report showing how many sheets include diagrams, pictures, or other drawing objects. | Validate a template by confirming required shapes are present on specific worksheets before further processing. | Filter and process only those worksheets that contain shapes when exporting to another format.
// AI Prompts: Generate C# code with Aspose.Cells that lists the names of all worksheets that have at least one shape. | Provide a snippet that counts worksheets with shapes and logs their indices to the console using Aspose.Cells for .NET. | Show how to copy only the worksheets containing shapes into a new workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells, iterates through each worksheet, checks the Shapes collection, increments a counter for sheets with one or more shapes, and outputs the total number of worksheets that contain shapes.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your actual file)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath); // uses the provided load constructor

        int worksheetsWithShapes = 0;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The Shapes collection is available on each worksheet.
            // If the count is greater than zero, this worksheet contains at least one shape.
            if (sheet.Shapes.Count > 0)
            {
                worksheetsWithShapes++;
            }
        }

        // Output the result
        Console.WriteLine($"Worksheets containing shapes: {worksheetsWithShapes}");
    }
}
