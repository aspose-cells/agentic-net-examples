// Title: C# – Recursively Ungroup Nested Shape Groups in Aspose.Cells Until Only Individual Shapes Remain
// Description: Demonstrates how to create a workbook, add rectangle and oval shapes, group them into nested GroupShape objects, and then flatten the hierarchy by repeatedly detecting and calling Ungroup on each GroupShape until the worksheet contains only single shapes. The workbook is saved after all groups are removed.
// Keywords: Aspose.Cells ungroup shapes | GroupShape Ungroup C# | flatten nested groups Aspose.Cells | recursive shape ungroup .NET | worksheet.Shapes ungroup | remove group shapes Aspose.Cells | C# Aspose.Cells shape handling | Aspose.Cells shape grouping | Ungroup all shapes Aspose.Cells | nested GroupShape recursion
// Common Searches: how to ungroup nested shape groups Aspose.Cells C# | recursive ungroup of GroupShape in worksheet | flatten shape groups in Aspose.Cells workbook | remove all group shapes from Excel sheet using Aspose.Cells | C# code to ungroup shapes until only individual objects remain
// Developer Intent: The developer needs to dissolve every GroupShape on a worksheet so that only individual shapes are left.
// Use Cases: Prepare a diagram for per‑shape formatting before exporting the workbook. | Extract shape‑level data after eliminating grouping created during import. | Simplify complex drawings to enable downstream processing that requires single‑shape objects.
// AI Prompts: Write C# code with Aspose.Cells that recursively ungroups every GroupShape in a worksheet until only single shapes remain. | Suggest an alternative to the break‑and‑restart loop for ungrouping nested groups without modifying the foreach enumeration. | Explain how to programmatically confirm that no GroupShape objects exist after the ungrouping routine.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeUngroupDemo
{
    // Demonstrates how to create a workbook, add rectangle and oval shapes, group them into nested GroupShape objects, and then flatten the hierarchy by repeatedly detecting and calling Ungroup on each GroupShape until the worksheet contains only single shapes. The workbook is saved after all groups are removed.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook(); // create a new workbook
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Sample shapes ----------
                // Add some shapes to the worksheet
                Shape rect1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 40);
                Shape rect2 = worksheet.Shapes.AddRectangle(6, 0, 2, 0, 80, 40);
                Shape oval1 = worksheet.Shapes.AddOval(10, 0, 2, 0, 80, 40);

                // Group the first two rectangles
                GroupShape groupLevel1 = worksheet.Shapes.Group(new Shape[] { rect1, rect2 });

                // Note: Aspose.Cells does not allow grouping an existing GroupShape with other shapes.
                // To demonstrate ungrouping, we will add the oval to the same group by grouping all three shapes together.
                // This creates a single group containing all three shapes.
                GroupShape groupAll = worksheet.Shapes.Group(new Shape[] { groupLevel1, oval1 });

                // ---------- Recursive ungrouping ----------
                // Continue ungrouping until no GroupShape remains in the collection
                bool ungroupedAny;
                do
                {
                    ungroupedAny = false;

                    // Iterate over the shapes collection.
                    // Because the collection changes when a group is ungrouped,
                    // we break after the first ungroup operation and restart the loop.
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        if (shape.IsGroup) // check if the shape is a group
                        {
                            GroupShape grp = (GroupShape)shape;
                            grp.Ungroup(); // ungroup the current group
                            ungroupedAny = true;
                            break; // restart enumeration
                        }
                    }
                } while (ungroupedAny);

                // ---------- Save the workbook ----------
                workbook.Save("UngroupedShapesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
