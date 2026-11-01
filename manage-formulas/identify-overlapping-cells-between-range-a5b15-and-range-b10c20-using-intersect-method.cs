// Title: Find overlapping cells of two Excel ranges (A5:B15 and B10:C20) using Aspose.Cells Intersect in C#
// Common Searches: Aspose.Cells C# intersect two ranges to get common cells | How to retrieve overlapping cells between A5:B15 and B10:C20 with Aspose.Cells | C# example for using Range.Intersect in Aspose.Cells | Iterate cells of an intersected range and read values in Aspose.Cells .NET | Save workbook after extracting intersected cells using Aspose.Cells
// Tags: range intersect method Aspose.Cells | overlapping cells extraction Aspose.Cells | enumerate intersected cells .NET | save intersected range workbook Aspose.Cells | Excel range intersection example C#

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, defines two ranges (A5:B15 and B10:C20) on the first worksheet, obtains their intersected range with the Range.Intersect method, checks for a non‑empty result, iterates through each intersected cell to output its address and value, and finally saves the workbook as IntersectResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the two ranges using Aspose.Cells.Range
            Aspose.Cells.Range rangeA = sheet.Cells.CreateRange("A5", "B15");
            Aspose.Cells.Range rangeB = sheet.Cells.CreateRange("B10", "C20");

            // Get the intersected range
            Aspose.Cells.Range intersectRange = rangeA.Intersect(rangeB);

            // Check if there is an intersection
            if (intersectRange != null && intersectRange.RowCount > 0 && intersectRange.ColumnCount > 0)
            {
                Console.WriteLine("Overlapping cells:");
                // Iterate through the intersected cells
                for (int i = 0; i < intersectRange.RowCount; i++)
                {
                    for (int j = 0; j < intersectRange.ColumnCount; j++)
                    {
                        Cell cell = intersectRange[i, j];
                        Console.WriteLine($"{cell.Name} = {cell.Value}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No overlapping cells found.");
            }

            // Save the workbook
            string outputPath = "IntersectResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
