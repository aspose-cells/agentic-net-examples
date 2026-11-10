// Title: How to find overlapping cells between ranges A5:B15 and B10:C20 using the Range.Intersect method in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find intersecting cells between A5:B15 and B10:C20 | using Range.Intersect to get overlapping cells in an Excel worksheet with Aspose.Cells | C# example to list cell addresses that are common to two ranges in Aspose.Cells | retrieve intersect range cells and print their names with Aspose.Cells .NET | determine overlapping area of two Excel ranges programmatically using Aspose.Cells
// Tags: range intersect Aspose.Cells | retrieve intersecting cells C# | excel range overlap detection Aspose | iterate intersect range cells | Aspose.Cells workbook save after intersection

using System;
using Aspose.Cells;

// The sample creates a workbook, defines two ranges (A5:B15 and B10:C20), uses Range.Intersect to obtain the overlapping area, iterates through the intersecting cells to output their addresses, and saves the result as IntersectResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the two ranges to compare (use Aspose.Cells.Range explicitly)
            Aspose.Cells.Range rangeA = sheet.Cells.CreateRange("A5:B15");
            Aspose.Cells.Range rangeB = sheet.Cells.CreateRange("B10:C20");

            // Get the intersecting range
            Aspose.Cells.Range intersectRange = rangeA.Intersect(rangeB);

            // Check if there is an intersection
            if (intersectRange != null && intersectRange.RowCount > 0 && intersectRange.ColumnCount > 0)
            {
                Console.WriteLine("Intersecting cells:");
                // Iterate through each cell in the intersecting range
                for (int i = 0; i < intersectRange.RowCount; i++)
                {
                    for (int j = 0; j < intersectRange.ColumnCount; j++)
                    {
                        Cell cell = intersectRange[i, j];
                        Console.WriteLine(cell.Name);
                    }
                }
            }
            else
            {
                Console.WriteLine("No overlapping cells found.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("IntersectResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
