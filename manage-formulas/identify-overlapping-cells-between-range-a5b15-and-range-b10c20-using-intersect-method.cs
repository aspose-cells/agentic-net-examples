// Title: Aspose.Cells C# Intersect Method – Find & Highlight Overlapping Cells (A5:B15 & B10:C20)
// Description: Creates a workbook, defines ranges A5:B15 and B10:C20, uses the Intersect method to retrieve their common cells, checks for null, applies a yellow background style to the intersected area, and saves the file as IntersectRangesDemo.xlsx.
// Keywords: Aspose.Cells Intersect C# | overlapping ranges Excel | highlight intersected cells | range intersection Aspose | apply style to range | Excel cell range intersect example
// Common Searches: Aspose.Cells intersect two ranges C# | How to highlight overlapping cells in Aspose.Cells | C# get common cells between A5:B15 and B10:C20 | Aspose.Cells range intersection null check | Apply background color to intersected range Aspose
// Developer Intent: Identify the cells that are common to two defined ranges and apply a visual highlight to them.
// Use Cases: Show the intersecting area of two data blocks by coloring the overlapping cells. | Validate range overlap before performing calculations such as summing only the shared cells. | Log or display the address of the intersected range for reporting or user feedback.
// AI Prompts: Generate C# code with Aspose.Cells that finds the intersection of ranges A5:B15 and B10:C20, colors the intersected cells yellow, and saves the workbook. | Explain the Intersect method in Aspose.Cells, including how to handle a null result before styling the range. | Provide step‑by‑step instructions to retrieve the address of an intersected range and output it to the console.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, defines ranges A5:B15 and B10:C20, uses the Intersect method to retrieve their common cells, checks for null, applies a yellow background style to the intersected area, and saves the file as IntersectRangesDemo.xlsx.
    public class IntersectRangesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create the first range A5:B15
                AsposeRange range1 = cells.CreateRange("A5", "B15");

                // Create the second range B10:C20
                AsposeRange range2 = cells.CreateRange("B10", "C20");

                // Get the rectangular intersection of the two ranges
                AsposeRange intersectRange = range1.Intersect(range2);

                // If an intersection exists, display its address and highlight it
                if (intersectRange != null)
                {
                    Console.WriteLine("Intersected range: " + intersectRange.Address);

                    // Apply a yellow background to the intersected cells
                    Style highlight = workbook.CreateStyle();
                    highlight.ForegroundColor = Color.Yellow;
                    highlight.Pattern = BackgroundType.Solid;
                    intersectRange.SetStyle(highlight);
                }
                else
                {
                    Console.WriteLine("The specified ranges do not intersect.");
                }

                // Save the workbook with the highlighted intersection
                workbook.Save("IntersectRangesDemo.xlsx");
                Console.WriteLine("Workbook saved as IntersectRangesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            IntersectRangesDemo.Run();
        }
    }
}
