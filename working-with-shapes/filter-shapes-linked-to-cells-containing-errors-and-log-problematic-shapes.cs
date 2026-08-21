// Title: Detect and Log Shapes Linked to Error Cells with Aspose.Cells for .NET
// Description: Loads an Excel file using LoadOptions.IgnoreUselessShapes, scans each worksheet's ShapeCollection, retrieves each shape's linked cell via GetLinkedCell, checks if the cell contains an error (IsErrorValue), and writes the shape name, type and error details to the console before optionally saving the workbook.
// Keywords: Aspose.Cells shape error detection | C# GetLinkedCell error cell | filter shapes by linked cell value | IgnoreUselessShapes performance | log shapes referencing #N/A or #DIV/0! | Excel shape linked cell validation
// Common Searches: Aspose.Cells find shapes linked to error cells | C# list shapes with #REF! reference | GetLinkedCell returns error cell Aspose | ignore useless shapes when loading workbook | log shape name and type for error cells
// Developer Intent: Identify shapes that reference cells with error values and record their identifiers for review or cleanup.
// Use Cases: Generate a validation report of visual objects pointing to invalid formula results. | Automate removal or reassignment of shapes linked to error cells before publishing a workbook. | Audit workbooks to ensure no graphics are tied to cells showing #DIV/0!, #N/A, or other errors.
// AI Prompts: Write C# code using Aspose.Cells that deletes shapes whose linked cells contain any error value. | Show how to collect shape IDs linked to error cells into a List<int> for further processing. | Explain how to extend the sample to capture the specific Excel error type (e.g., #VALUE!, #REF!) from the linked cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeErrorFilter
{
    // Loads an Excel file using LoadOptions.IgnoreUselessShapes, scans each worksheet's ShapeCollection, retrieves each shape's linked cell via GetLinkedCell, checks if the cell contains an error (IsErrorValue), and writes the shape name, type and error details to the console before optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Load the workbook with options (ignore useless shapes to speed up loading)
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.IgnoreUselessShapes = true;
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Examine each shape in the worksheet
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Try to obtain the linked cell address (if any)
                    string linkedCellAddress = shape.GetLinkedCell(false, false);

                    // If the shape is linked to a cell, check whether that cell contains an error
                    if (!string.IsNullOrEmpty(linkedCellAddress))
                    {
                        Cell linkedCell = sheet.Cells[linkedCellAddress];

                        if (linkedCell.IsErrorValue)
                        {
                            // Log the problematic shape information
                            Console.WriteLine($"Shape '{shape.Name}' (Type={shape.Type}) is linked to error cell '{linkedCellAddress}' with value '{linkedCell.StringValue}'.");
                        }
                    }
                }
            }

            // Save the workbook (unchanged, but could be a different file if needed)
            workbook.Save("output.xlsx");
        }
    }
}
