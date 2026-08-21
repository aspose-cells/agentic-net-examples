// Title: C# – Recursively Ungroup Nested Shape Groups in an Aspose.Cells Worksheet
// Description: Demonstrates how to flatten a worksheet by recursively ungrouping every GroupShape in Aspose.Cells. The sample creates nested groups, then uses a safe backward‑iteration loop to call ShapeCollection.Ungroup until only individual shapes remain, and finally saves the workbook.
// Keywords: Aspose.Cells ungroup shapes | C# recursive shape ungroup | GroupShape flatten Excel | ShapeCollection.Ungroup method | nested shape groups Aspose.Cells | Excel shape manipulation .NET | remove grouped objects worksheet | flatten diagram Aspose.Cells | programmatic shape ungrouping | Aspose.Cells shape collection
// Common Searches: how to ungroup all nested shape groups in Aspose.Cells | recursive ungrouping of GroupShape objects C# | flatten grouped shapes in Excel using Aspose.Cells | remove shape groups from worksheet programmatically | Aspose.Cells ShapeCollection Ungroup example
// Developer Intent: Programmatically ungroup every nested GroupShape in a worksheet so that only individual shapes remain.
// Use Cases: Prepare a workbook for PDF or image export by flattening complex diagrams. | Enable shape‑level formatting after receiving an Excel file with grouped objects. | Clean up imported spreadsheets that contain nested groups before further data processing.
// AI Prompts: Generate a C# method that recursively ungroups all GroupShape objects in an Aspose.Cells worksheet, handling collection changes safely. | Show code that detects nested shape groups, flattens them, and saves the workbook with only individual shapes. | Explain how to modify the UngroupAllNestedGroups method to return a list of shapes that were extracted from groups.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeUngroupDemo
{
    // Demonstrates how to flatten a worksheet by recursively ungrouping every GroupShape in Aspose.Cells. The sample creates nested groups, then uses a safe backward‑iteration loop to call ShapeCollection.Ungroup until only individual shapes remain, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Sample data: create nested groups for demonstration purposes
                // ------------------------------------------------------------
                // Add some basic shapes
                Shape rect1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 40);
                Shape rect2 = worksheet.Shapes.AddRectangle(6, 0, 2, 0, 80, 40);
                Shape oval1 = worksheet.Shapes.AddOval(10, 0, 2, 0, 80, 40);

                // First level group (rect1 + rect2)
                GroupShape groupLevel1 = worksheet.Shapes.Group(new Shape[] { rect1, rect2 });

                // Second level group: include oval1 and nest the first group inside it
                GroupShape groupLevel2 = worksheet.Shapes.Group(new Shape[] { oval1, groupLevel1 });

                // ------------------------------------------------------------
                // Recursive ungrouping: flatten all nested groups
                // ------------------------------------------------------------
                UngroupAllNestedGroups(worksheet);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("UngroupedShapesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        /// <param name="worksheet">Target worksheet.</param>
        static void UngroupAllNestedGroups(Worksheet worksheet)
        {
            ShapeCollection shapes = worksheet.Shapes;

            // Continue looping while there is at least one group shape present
            bool groupsRemaining = true;
            while (groupsRemaining)
            {
                groupsRemaining = false;

                // Iterate backwards because ungrouping modifies the collection
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    Shape shape = shapes[i];

                    // Identify group shapes using IsGroup property
                    if (shape.IsGroup)
                    {
                        // Ungroup the current group (rule: ShapeCollection.Ungroup)
                        shapes.Ungroup((GroupShape)shape);
                        // After ungrouping, there may be new groups, so we set the flag
                        groupsRemaining = true;
                    }
                }
            }
        }
    }
}
