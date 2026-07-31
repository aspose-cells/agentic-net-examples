// Title: C# Aspose.Cells – Find and Highlight Overlapping Cells Using Range.Intersect
// Description: Creates a workbook, defines ranges A5:B15 and B10:C20, retrieves their intersection with the Range.Intersect method, displays the intersected address, applies a light‑green fill, and saves the file as IntersectResult.xlsx.
// Keywords: Aspose.Cells Range.Intersect C# | overlapping cells Excel | highlight intersected range | C# Excel inter‑section example | Aspose.Cells formatting intersect
// Common Searches: Aspose.Cells intersect two ranges C# | how to get overlapping cells A5:B15 B10:C20 | highlight intersected cells Aspose.Cells | Range.Intersect null check example
// Developer Intent: Identify the common cells of two specified ranges and apply visual formatting to them.
// Use Cases: Apply conditional formatting to the area where two data tables overlap. | Validate user‑defined zones before merging worksheets to avoid data collisions. | Generate a visual cue for intersecting ranges in automated report generation.
// AI Prompts: Generate C# code with Aspose.Cells that finds the intersection of A5:B15 and B10:C20 and colors the cells yellow. | Explain the behavior of Range.Intersect in Aspose.Cells and how to handle a null return safely. | Show an alternative way to style intersected cells using style inheritance instead of SetStyle.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsIntersectDemo
{
    // Creates a workbook, defines ranges A5:B15 and B10:C20, retrieves their intersection with the Range.Intersect method, displays the intersected address, applies a light‑green fill, and saves the file as IntersectResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data (optional, just for visualization)
                for (int row = 4; row <= 19; row++) // rows 5 to 20 (zero‑based index)
                {
                    for (int col = 0; col <= 2; col++) // columns A (0) to C (2)
                    {
                        cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Create the two ranges to be intersected
                // Range A5:B15  => rows 4‑14, columns 0‑1
                AsposeRange rangeA = cells.CreateRange("A5", "B15");
                // Range B10:C20 => rows 9‑19, columns 1‑2
                AsposeRange rangeB = cells.CreateRange("B10", "C20");

                // Use the Intersect method to get the overlapping area
                AsposeRange intersectRange = rangeA.Intersect(rangeB);

                if (intersectRange != null)
                {
                    // Output the address of the intersected range
                    Console.WriteLine("Intersected range address: " + intersectRange.Address);

                    // Highlight the intersected cells with a background color
                    Style highlight = workbook.CreateStyle();
                    highlight.ForegroundColor = Color.LightGreen;
                    highlight.Pattern = BackgroundType.Solid;
                    intersectRange.SetStyle(highlight);
                }
                else
                {
                    Console.WriteLine("The ranges do not intersect.");
                }

                // Save the workbook to verify the result
                workbook.Save("IntersectResult.xlsx");
                Console.WriteLine("Workbook saved as IntersectResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
