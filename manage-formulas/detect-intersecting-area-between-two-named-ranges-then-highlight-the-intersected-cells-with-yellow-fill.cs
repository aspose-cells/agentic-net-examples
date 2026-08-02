// Title: C# – Highlight Intersection of Two Named Ranges with Yellow Fill using Aspose.Cells
// Description: Creates a workbook, defines two named ranges (A1:C3 and B2:D4), computes their overlapping area with the Range.Intersect method, applies a solid yellow background to the intersected cells, and saves the result as an Excel file.
// Keywords: Aspose.Cells C# intersect named ranges | Range.Intersect method | highlight overlapping cells | apply yellow fill style | Excel workbook styling Aspose
// Common Searches: Aspose.Cells find intersecting cells between named ranges | C# highlight overlapping area in Excel with Aspose | Range.Intersect usage example | apply background color to intersected range Aspose.Cells | how to style intersected cells in .NET Excel library
// Developer Intent: Locate the cells where two named ranges overlap and color that region yellow.
// Use Cases: Visually mark the common area of two data blocks for quick analysis. | Validate range overlap before performing calculations that depend on shared cells. | Generate reports that automatically highlight intersecting sections of named ranges.
// AI Prompts: Generate C# code with Aspose.Cells that finds the intersection of two named ranges and fills the intersected cells with yellow. | Show how to retrieve Range objects from Name objects and use Range.Intersect to style the overlapping area. | Explain step‑by‑step how to create named ranges, compute their intersection, and apply a solid fill color in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace IntersectNamedRangesDemo
{
    // Creates a workbook, defines two named ranges (A1:C3 and B2:D4), computes their overlapping area with the Range.Intersect method, applies a solid yellow background to the intersected cells, and saves the result as an Excel file.
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

                // Populate sample data
                for (int i = 0; i < 5; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        cells[i, j].PutValue($"R{i}C{j}");
                    }
                }

                // Define first named range: A1:C3
                int idx1 = workbook.Worksheets.Names.Add("FirstRange");
                Name name1 = workbook.Worksheets.Names[idx1];
                name1.RefersTo = "=Sheet1!$A$1:$C$3";

                // Define second named range: B2:D4
                int idx2 = workbook.Worksheets.Names.Add("SecondRange");
                Name name2 = workbook.Worksheets.Names[idx2];
                name2.RefersTo = "=Sheet1!$B$2:$D$4";

                // Retrieve the actual Range objects from the names
                AsposeRange[] ranges1 = name1.GetRanges(); // should contain one range
                AsposeRange[] ranges2 = name2.GetRanges();

                if (ranges1.Length == 0 || ranges2.Length == 0)
                {
                    Console.WriteLine("One of the named ranges does not contain any cells.");
                    return;
                }

                AsposeRange rangeA = ranges1[0];
                AsposeRange rangeB = ranges2[0];

                // Compute the intersection of the two ranges
                AsposeRange intersect = rangeA.Intersect(rangeB);

                if (intersect != null)
                {
                    // Create a style with yellow fill
                    Style yellowStyle = workbook.CreateStyle();
                    yellowStyle.ForegroundColor = Color.Yellow;
                    yellowStyle.Pattern = BackgroundType.Solid;

                    // Apply the style to the intersected area
                    intersect.SetStyle(yellowStyle);

                    Console.WriteLine($"Intersection found: {intersect.Address}");
                }
                else
                {
                    Console.WriteLine("The named ranges do not intersect.");
                }

                // Save the workbook
                string outputPath = "IntersectedNamedRanges.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
