// Title: Lock a grouped picture and chart in Aspose.Cells for .NET and confirm that ungrouping is blocked
// AI Prompts: Write C# code with Aspose.Cells that inserts a PNG image and a column chart, groups them into a GroupShape, sets the group's IsLocked property to true, then attempts to ungroup and captures the resulting exception. | Show how to test lock enforcement on a GroupShape containing a picture and a chart by calling Ungroup and handling the expected error in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# lock grouped shapes to prevent ungrouping | how to enforce IsLocked on a GroupShape in Aspose.Cells .NET | example of grouping a picture and a chart and testing lock in Aspose.Cells | Ungroup throws exception when GroupShape IsLocked is true Aspose.Cells | C# Aspose.Cells group shapes and verify lock enforcement
// Tags: group shapes with Aspose.Cells | lock GroupShape in Aspose.Cells | prevent ungrouping Aspose.Cells .NET | picture and chart grouping Aspose.Cells | IsLocked property Aspose.Cells | handle Ungroup exception Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// Creates a workbook, adds a PNG picture and a column chart, groups them, locks the GroupShape, attempts to ungroup to verify the lock, and saves the file as GroupedLocked.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a picture to the worksheet if the file exists
            Shape? pictureShape = null;
            string imagePath = "sample.png";
            if (File.Exists(imagePath))
            {
                try
                {
                    // Add returns the index of the picture; retrieve the Picture object
                    int picIdx = sheet.Pictures.Add(1, 1, imagePath);
                    pictureShape = sheet.Pictures[picIdx];
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add picture: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Add a chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 1, 15, 10);
            Chart chart = sheet.Charts[chartIdx];
            // The chart itself is a shape; obtain its ChartObject (inherits from Shape)
            Shape chartShape = chart.ChartObject;

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);
            chart.NSeries.Add("A1:B2", true);

            // Group the picture and the chart together (only if picture was added)
            if (pictureShape != null)
            {
                try
                {
                    Shape[] shapesToGroup = new Shape[] { pictureShape, chartShape };
                    GroupShape group = sheet.Shapes.Group(shapesToGroup);

                    // Lock the group to prevent modifications
                    group.IsLocked = true;

                    // Attempt to ungroup the locked group to test lock enforcement
                    try
                    {
                        sheet.Shapes.Ungroup(group);
                        Console.WriteLine("Ungroup succeeded (lock not enforced).");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Ungroup failed as expected: " + ex.Message);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during grouping: {ex.Message}");
                }
            }

            // Save the workbook
            string outputPath = "GroupedLocked.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
