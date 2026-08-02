// Title: Lock a GroupShape of a picture and chart and verify ungroup protection with Aspose.Cells for .NET
// Description: Demonstrates how to add an image and a column chart to a worksheet, group them, lock the GroupShape (including the Ungroup property), protect the sheet, and attempt to ungroup to confirm that the lock is enforced. The workbook is saved as "GroupLockDemo.xlsx".
// Keywords: Aspose.Cells group shape lock | lock grouped picture chart .NET | prevent ungroup Aspose.Cells | worksheet protection shape lock C# | GroupShape Ungroup restriction
// Common Searches: how to lock a grouped shape in Aspose.Cells | prevent ungrouping of chart and image Aspose.Cells C# | group picture and chart then protect worksheet Aspose.Cells | test shape lock enforcement Aspose.Cells .NET | lock GroupShape Ungroup property
// Developer Intent: Group a picture and a chart, lock the group, protect the worksheet, and confirm that ungrouping is blocked.
// Use Cases: Create a read‑only dashboard where visual elements stay together. | Distribute Excel templates that keep grouped objects intact while allowing cell edits. | Validate shape‑locking behavior by programmatically attempting an ungroup operation.
// AI Prompts: Show C# code to lock a GroupShape and disable its Ungroup operation using Aspose.Cells. | Explain how to protect a worksheet while keeping grouped shapes locked for end users. | Provide a method to check the lock status of a GroupShape before calling Ungroup.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsGroupLockDemo
{
    // Demonstrates how to add an image and a column chart to a worksheet, group them, lock the GroupShape (including the Ungroup property), protect the sheet, and attempt to ungroup to confirm that the lock is enforced. The workbook is saved as "GroupLockDemo.xlsx".
    public class Program
    {
        public static void Main()
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
                    int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
                    picture = worksheet.Pictures[pictureIndex];
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Add a chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                Shape chartShape = chart.ChartObject;

                // Ensure we have both shapes before grouping
                if (picture != null)
                {
                    // Group the picture and the chart shape together
                    GroupShape groupShape = worksheet.Shapes.Group(new Shape[] { picture, chartShape });

                    // Lock the group shape to prevent modifications when the sheet is protected
                    groupShape.IsLocked = true;
                    // Additionally lock the ungroup operation specifically
                    groupShape.SetLockedProperty(ShapeLockType.Ungroup, true);

                    // Protect the worksheet so that locked objects cannot be edited
                    worksheet.Protect(ProtectionType.All);

                    // Attempt to ungroup the locked group shape and capture the result
                    try
                    {
                        groupShape.Ungroup();
                        Console.WriteLine("Ungroup succeeded (unexpected).");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Ungroup failed as expected: " + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("Picture not added; skipping grouping and locking operations.");
                }

                // Save the workbook
                workbook.Save("GroupLockDemo.xlsx");
                Console.WriteLine("Workbook saved as 'GroupLockDemo.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
