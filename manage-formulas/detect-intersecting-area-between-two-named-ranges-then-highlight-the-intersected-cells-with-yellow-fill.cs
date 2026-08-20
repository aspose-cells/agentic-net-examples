// Title: Aspose.Cells .NET – Highlight Intersection of Two Named Ranges with Yellow Fill
// Description: Creates a workbook, defines two named ranges (A1:B3 and B2:C4), uses the Range.Intersect method to locate overlapping cells, applies a solid yellow background style to the intersected area, and saves the result as an XLSX file.
// Keywords: Aspose.Cells intersect named ranges | Range.Intersect C# | highlight overlapping cells Aspose | apply yellow fill Aspose.Cells | C# Excel range styling | named range intersection example
// Common Searches: Aspose.Cells find intersection of named ranges | C# highlight intersecting cells Excel | how to use Range.Intersect in Aspose.Cells | apply background color to intersected range Aspose | example code for named range overlap .NET
// Developer Intent: Detect overlapping cells of two named ranges and color them yellow.
// Use Cases: Visualize data overlap by shading the intersecting region of two named ranges. | Validate range relationships before performing calculations or reporting. | Generate formatted Excel reports where intersected cells need emphasis.
// AI Prompts: Generate C# code with Aspose.Cells that retrieves two named ranges, computes their intersection, and fills the intersected cells with a yellow background. | Provide a robust Aspose.Cells example that checks for a null intersection before applying any style and saves the workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace IntersectNamedRangesDemo
{
    // Creates a workbook, defines two named ranges (A1:B3 and B2:C4), uses the Range.Intersect method to locate overlapping cells, applies a solid yellow background style to the intersected area, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample data
                cells["A1"].PutValue("A1");
                cells["B1"].PutValue("B1");
                cells["C1"].PutValue("C1");
                cells["A2"].PutValue("A2");
                cells["B2"].PutValue("B2");
                cells["C2"].PutValue("C2");
                cells["A3"].PutValue("A3");
                cells["B3"].PutValue("B3");
                cells["C3"].PutValue("C3");

                // Define two named ranges
                // Range1: A1:B3
                int idx1 = workbook.Worksheets.Names.Add("Range1");
                Name name1 = workbook.Worksheets.Names[idx1];
                name1.RefersTo = "=Sheet1!$A$1:$B$3";

                // Range2: B2:C4 (C4 will be empty but still part of the range)
                int idx2 = workbook.Worksheets.Names.Add("Range2");
                Name name2 = workbook.Worksheets.Names[idx2];
                name2.RefersTo = "=Sheet1!$B$2:$C$4";

                // Retrieve the actual Range objects from the names
                AsposeRange[] ranges1 = name1.GetRanges(); // should contain one range
                AsposeRange[] ranges2 = name2.GetRanges();

                // Ensure both named ranges have at least one range defined
                if (ranges1.Length == 0 || ranges2.Length == 0)
                {
                    Console.WriteLine("One of the named ranges does not refer to a valid range.");
                    return;
                }

                // Get the first (and only) range from each name
                AsposeRange rangeA = ranges1[0];
                AsposeRange rangeB = ranges2[0];

                // Compute the intersection of the two ranges
                AsposeRange intersectRange = rangeA.Intersect(rangeB);

                if (intersectRange != null)
                {
                    // Create a style with yellow background
                    Style yellowStyle = workbook.CreateStyle();
                    yellowStyle.ForegroundColor = Color.Yellow;
                    yellowStyle.Pattern = BackgroundType.Solid;

                    // Apply the style to the intersected cells
                    intersectRange.SetStyle(yellowStyle);

                    Console.WriteLine($"Intersection found: {intersectRange.Address}");
                }
                else
                {
                    Console.WriteLine("The named ranges do not intersect.");
                }

                // Save the workbook
                string outputPath = "IntersectNamedRangesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
