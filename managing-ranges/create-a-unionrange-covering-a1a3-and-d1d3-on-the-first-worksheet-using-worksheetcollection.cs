// Title: Aspose.Cells C# – Create UnionRange A1:A3,D1:D3 on First Worksheet
// Description: Shows how to call WorksheetCollection.CreateUnionRange in Aspose.Cells for .NET to build a UnionRange that includes A1:A3 and D1:D3 on the first worksheet, assign a single value, and save the workbook.
// Keywords: Aspose.Cells | C# | UnionRange | CreateUnionRange | WorksheetCollection | non‑contiguous range | set value | save workbook | Excel automation | in‑memory workbook
// Common Searches: Aspose.Cells create union range C# | WorksheetCollection CreateUnionRange example | set value for multiple cells Aspose.Cells | non adjacent cells union range .NET | save workbook after union range
// Developer Intent: Create a UnionRange that spans A1:A3 and D1:D3 on the first worksheet and write a single value to the combined area.
// Use Cases: Apply identical formatting to two separate cell blocks with one command. | Insert the same header or label into disjoint locations without looping. | Assign a constant value or formula to multiple non‑adjacent ranges for consistent data entry.
// AI Prompts: Write C# code that creates a UnionRange for A1:A3 and D1:D3, applies a formula, and saves the file. | Demonstrate how to set a cell style on a UnionRange created via WorksheetCollection.CreateUnionRange in Aspose.Cells. | Explain how to enumerate each area of a UnionRange and process its cells using Aspose.Cells for .NET.

using Aspose.Cells;

// Shows how to call WorksheetCollection.CreateUnionRange in Aspose.Cells for .NET to build a UnionRange that includes A1:A3 and D1:D3 on the first worksheet, assign a single value, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (in‑memory)
        Workbook workbook = new Workbook();

        // Index of the first worksheet (0‑based)
        int sheetIndex = 0;

        // Create a UnionRange that includes A1:A3 and D1:D3 on the first worksheet
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:A3,D1:D3", sheetIndex);

        // Example: set a value for the whole union range to verify it works
        unionRange.Value = "Test";

        // Save the workbook to a file
        workbook.Save("UnionRangeOutput.xlsx");
    }
}
