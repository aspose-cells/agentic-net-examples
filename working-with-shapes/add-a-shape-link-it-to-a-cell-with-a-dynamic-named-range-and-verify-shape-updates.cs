// Title: Add a rectangle shape linked to a dynamic OFFSET named range and reposition it as the range expands using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a rectangle shape, anchors it to the first cell of a dynamic named range defined with OFFSET, and saves the workbook. | Show how to add more rows to the source column, recalculate formulas, and move the existing shape to the new start cell of the named range. | Provide code that outputs the shape's row and column coordinates before and after the range expansion to verify the update.
// Common Searches: how to attach a shape to a dynamic OFFSET named range in Aspose.Cells C# | Aspose.Cells move rectangle shape when named range size changes | using freefloating placement for shapes linked to cells Aspose.Cells | recalculate named range after adding rows Aspose.Cells C# | save workbook with shape linked to dynamic range Aspose.Cells
// Tags: dynamic named range shape linking Aspose.Cells | rectangle shape anchored to cell C# | freefloating shape placement Aspose.Cells | recalculate formulas after data addition Aspose.Cells | update shape position with expanded OFFSET range

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // The example creates a workbook, defines a dynamic named range using an OFFSET formula, adds a rectangle shape anchored to the range's first cell, saves the file, appends additional rows to expand the range, recalculates formulas, moves the shape to the new start cell, and saves the updated workbook while printing the shape's coordinates before and after the change.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate column A with some initial data
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                }

                // Define a dynamic named range "MyRange" that expands with data in column A
                // Formula: =OFFSET(DataSheet!$A$1,0,0,COUNTA(DataSheet!$A:$A),1)
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRange = workbook.Worksheets.Names[nameIndex];
                myRange.RefersTo = "=OFFSET(DataSheet!$A$1,0,0,COUNTA(DataSheet!$A:$A),1)";

                // Retrieve the first cell of the named range to set the shape position
                Aspose.Cells.Range range = sheet.Cells.CreateRange(myRange.RefersTo);
                int startRow = range.FirstRow;
                int startColumn = range.FirstColumn;

                // Add a rectangle shape anchored to the start cell
                Shape rectShape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, startRow, startColumn, 0, 0, 100, 30);
                rectShape.Name = "LinkedRectangle";
                rectShape.Placement = PlacementType.FreeFloating; // Allows movement with cells

                // Set some text inside the shape to identify it
                rectShape.Text = "Linked to MyRange";

                // Save the workbook after initial setup
                workbook.Save("ShapeLinkedToDynamicRange_Initial.xlsx");

                // Verify initial linking: shape should be anchored to A1 (first cell of the range)
                Console.WriteLine($"Initial Shape Position -> Row: {rectShape.UpperLeftRow}, Column: {rectShape.UpperLeftColumn}");

                // Add more data to column A to expand the dynamic named range
                for (int i = 5; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                }

                // Recalculate formulas so that the named range updates
                workbook.CalculateFormula();

                // Retrieve the updated range (it should now include rows 0-9 in column A)
                Aspose.Cells.Range updatedRange = sheet.Cells.CreateRange(myRange.RefersTo);
                int updatedStartRow = updatedRange.FirstRow;
                int updatedStartColumn = updatedRange.FirstColumn;

                // Move the shape to the new start cell of the dynamic range
                rectShape.UpperLeftRow = updatedStartRow;
                rectShape.UpperLeftColumn = updatedStartColumn;

                // Save the workbook after updating
                workbook.Save("ShapeLinkedToDynamicRange_Updated.xlsx");

                // Verify that the shape's position reflects the updated range
                Console.WriteLine($"Updated Shape Position -> Row: {rectShape.UpperLeftRow}, Column: {rectShape.UpperLeftColumn}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
