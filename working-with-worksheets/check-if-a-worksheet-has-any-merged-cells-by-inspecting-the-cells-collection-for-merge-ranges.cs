// Title: How to detect merged cells in an Aspose.Cells worksheet (C#/.NET)
// Description: C# example that creates a workbook, merges cells A1:B2, calls Cells.GetMergedAreas() to retrieve all merged ranges, checks if any exist, prints each range, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | merged cells detection | GetMergedAreas | CellArea | worksheet merge detection | list merged ranges
// Common Searches: Aspose.Cells check for merged cells C# | GetMergedAreas example Aspose.Cells | how to find merged cells in Excel using Aspose.Cells | C# detect merged ranges in worksheet | list merged cells Aspose.Cells .NET
// Developer Intent: Determine whether a worksheet contains any merged cells and retrieve their ranges.
// Use Cases: Validate a worksheet before PDF export to avoid layout issues caused by merged cells. | Iterate over merged areas to apply custom formatting or data validation. | Conditionally unmerge cells based on business rules after detection.
// AI Prompts: Write a C# method using Aspose.Cells that returns true if a worksheet has merged cells, otherwise false. | Generate code that extracts all merged cell ranges from a worksheet and returns them as a list of strings. | Provide an example that logs each merged area and then removes all merged cells from the worksheet.

using System;
using Aspose.Cells;

// C# example that creates a workbook, merges cells A1:B2, calls Cells.GetMergedAreas() to retrieve all merged ranges, checks if any exist, prints each range, and saves the file.
public class CheckMergedCellsDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Example merge: A1:B2 (rows 0-1, columns 0-1)
        cells.Merge(0, 0, 2, 2);

        // Retrieve all merged areas in the worksheet
        CellArea[] mergedAreas = cells.GetMergedAreas();

        // Determine whether any merged cells exist
        bool hasMergedCells = mergedAreas != null && mergedAreas.Length > 0;
        Console.WriteLine("Worksheet has merged cells? " + hasMergedCells);

        // If there are merged cells, output their ranges
        if (hasMergedCells)
        {
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine($"Merged area: Row[{area.StartRow}-{area.EndRow}], Column[{area.StartColumn}-{area.EndColumn}]");
            }
        }

        // Save the workbook (optional)
        workbook.Save("CheckMergedCellsDemo.xlsx");
    }
}
