// Title: Find overlapping cells of ranges A5:B15 and B10:C20 using Aspose.Cells Range.Intersect in C#
// AI Prompts: Generate C# code that creates two Aspose.Cells ranges (A5:B15 and B10:C20) and prints the addresses of their intersecting cells. | Write a reusable method in C# that accepts two Aspose.Cells Range objects and returns a list of cell names representing their intersection. | Show how to save the workbook after retrieving intersected cells with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# how to get cells that overlap between two ranges | example of Range.Intersect method for Excel ranges in Aspose.Cells | C# code to list intersecting cells of A5:B15 and B10:C20 using Aspose.Cells | retrieve overlapping cell addresses from two ranges with Aspose.Cells library | using CreateRange and Intersect to find common cells in an Excel worksheet C#
// Tags: Aspose.Cells range intersect C# | extract overlapping cell addresses Aspose.Cells | CreateRange and Intersect usage example | list intersected cells in Excel workbook | save workbook after range intersection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, defines two ranges A5:B15 and B10:C20 on the first worksheet, uses the Range.Intersect method to obtain any overlapping cells, prints each intersecting cell's address, and saves the workbook as IntersectExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the two ranges
            AsposeRange rangeA = sheet.Cells.CreateRange("A5", "B15");
            AsposeRange rangeB = sheet.Cells.CreateRange("B10", "C20");

            // Find the intersecting range
            AsposeRange intersect = rangeA.Intersect(rangeB);

            // Output the addresses of the overlapping cells, if any
            if (intersect != null)
            {
                foreach (Cell cell in intersect)
                {
                    Console.WriteLine(cell.Name);
                }
            }
            else
            {
                Console.WriteLine("No overlapping cells found.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle handling)
            string outputPath = "IntersectExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
