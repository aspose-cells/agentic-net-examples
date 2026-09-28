// Title: Highlight the intersecting cells of two named ranges with a yellow fill using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, retrieve the named ranges "Range1" and "Range2", compute their intersection with Range.Intersect, create a solid yellow style, and apply it to the intersected cells using Aspose.Cells. | In C#, get two named ranges from a workbook, determine the overlapping area, build a yellow background style, apply the style to the intersected range, and save the updated file.
// Common Searches: c# aspocells find intersection of two named ranges | how to apply yellow fill to overlapping cells in Excel using Aspose.Cells | Aspose.Cells .NET highlight intersecting area of named ranges | range.Intersect example in Aspose.Cells C# | retrieve named range by name and style intersected cells Aspose
// Tags: range.Intersect method Aspose.Cells | named range retrieval Aspose.Cells | apply solid fill style to range Aspose.Cells | highlight overlapping cells Excel C# | cell shading with Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads input.xlsx, obtains the named ranges "Range1" and "Range2", determines their intersecting area, applies a solid yellow background style to the intersected cells, and saves the result as output.xlsx.
class IntersectAndHighlight
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the two named ranges by their names
            // GetRangeByName returns an Aspose.Cells.Range object or null if not found
            AsposeRange range1 = workbook.Worksheets.GetRangeByName("Range1");
            AsposeRange range2 = workbook.Worksheets.GetRangeByName("Range2");

            if (range1 == null || range2 == null)
            {
                Console.WriteLine("One or both named ranges (Range1, Range2) were not found in the workbook.");
                return;
            }

            // Find the intersecting area between the two ranges
            AsposeRange intersectRange = range1.Intersect(range2);

            // If there is an intersection, apply a yellow fill to the intersected cells
            if (intersectRange != null && intersectRange.RowCount > 0 && intersectRange.ColumnCount > 0)
            {
                // Create a style with solid yellow background
                Style yellowStyle = workbook.CreateStyle();
                yellowStyle.ForegroundColor = Color.Yellow;
                yellowStyle.Pattern = BackgroundType.Solid;

                // Define which style attributes to apply (only cell shading)
                StyleFlag styleFlag = new StyleFlag { CellShading = true };

                // Apply the style to the intersected range
                intersectRange.ApplyStyle(yellowStyle, styleFlag);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
