// Title: Unmerge the specific cell range D4:F4 in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unmerge cells D4 to F4 in an existing Excel file | remove merged range D4:F4 using Aspose.Cells worksheet.Cells.UnMerge | how to restore original cell boundaries after merging with Aspose.Cells .NET | unmerge specific cells in Excel with Aspose.Cells C# code example | Aspose.Cells programmatically unmerge merged cells in a .NET workbook
// Tags: Aspose.Cells unmerge cell range C# | worksheet.Cells.UnMerge method example | restore original cell boundaries Excel .NET | remove merged cells D4:F4 Aspose.Cells | unmerge specific range Excel using Aspose.Cells

using Aspose.Cells;

// The sample loads an existing Excel workbook, accesses the first worksheet, calls worksheet.Cells.UnMerge(3, 3, 1, 3) to unmerge the previously merged range D4:F4, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify the appropriate index/name)
        var worksheet = workbook.Worksheets[0];

        // Unmerge the previously merged range D4:F4
        // D4 corresponds to row index 3 (zero‑based) and column index 3
        // The range spans 1 row and 3 columns (D, E, F)
        worksheet.Cells.UnMerge(3, 3, 1, 3);

        // Save the workbook with the changes (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
