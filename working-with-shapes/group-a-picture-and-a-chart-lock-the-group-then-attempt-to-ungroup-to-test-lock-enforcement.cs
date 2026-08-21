// Title: Aspose.Cells for .NET – Group a Picture and Chart, Lock the Group, Protect Worksheet, and Test Ungroup Blocking (C#)
// Description: Demonstrates how to add an image and a column chart to a worksheet, group them, apply a ShapeLockType.Ungroup lock, protect the sheet, attempt to ungroup, and verify that the lock is enforced before saving the workbook.
// Keywords: Aspose.Cells group shapes C# | lock grouped objects Aspose.Cells | ShapeLockType.Ungroup example | worksheet protection Aspose.Cells .NET | prevent ungrouping Excel C# | group picture and chart Aspose | Aspose.Cells shape locking tutorial | C# Excel shape group lock
// Common Searches: how to lock a grouped shape in Aspose.Cells .NET | prevent ungrouping of chart and image with Aspose.Cells | Aspose.Cells worksheet protection and shape locks | C# code to group picture and chart and lock them | test ShapeLockType.Ungroup enforcement
// Developer Intent: Create a grouped picture‑chart object, lock it against ungrouping, protect the worksheet, and confirm that the lock prevents the ungroup operation.
// Use Cases: Design a locked dashboard where a logo and its chart stay together for end‑users. | Distribute an Excel template that preserves layout by preventing shape separation. | Validate that ShapeLockType.Ungroup works correctly when worksheet protection is active.
// AI Prompts: Generate C# code using Aspose.Cells to group a picture and a chart, lock the group from ungrouping, protect the worksheet, and verify the lock. | Explain how ShapeLockType.Ungroup interacts with worksheet protection in Aspose.Cells for .NET. | Provide step‑by‑step instructions to test whether a locked group can be ungrouped and how to handle the exception.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsGroupLockDemo
{
    // Demonstrates how to add an image and a column chart to a worksheet, group them, apply a ShapeLockType.Ungroup lock, protect the sheet, attempt to ungroup, and verify that the lock is enforced before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a picture to the worksheet if the file exists
                Picture picture = null;
                string imagePath = "example.jpg";
                if (File.Exists(imagePath))
                {
                    int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                    picture = worksheet.Pictures[pictureIndex];
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture addition.");
                }

                // Add a chart to the worksheet (order of parameters may vary by Aspose.Cells version)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 2, 15, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Retrieve the shape that represents the chart
                Shape chartShape = worksheet.Shapes[chartIndex];

                // Group the picture and the chart shape (if picture was added)
                GroupShape group = null;
                try
                {
                    if (picture != null)
                    {
                        group = worksheet.Shapes.Group(new Shape[] { picture, chartShape });
                    }
                    else
                    {
                        group = worksheet.Shapes.Group(new Shape[] { chartShape });
                    }

                    // Lock the group against ungrouping
                    group.SetLockedProperty(ShapeLockType.Ungroup, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during grouping: {ex.Message}");
                    return;
                }

                // Protect the worksheet so that locked properties are enforced
                worksheet.Protection.AllowEditingObject = false; // disallow object editing
                worksheet.Protect(ProtectionType.All);

                // Attempt to ungroup the locked group
                try
                {
                    group.Ungroup();
                    Console.WriteLine("Ungroup operation completed.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during ungrouping: {ex.Message}");
                }

                // Verify whether the shapes are still in the group
                Console.WriteLine($"Picture IsInGroup: {(picture != null ? picture.IsInGroup.ToString() : "N/A")}");
                Console.WriteLine($"ChartShape IsInGroup: {chartShape.IsInGroup}");

                // Save the workbook
                string outputPath = "GroupLockTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
