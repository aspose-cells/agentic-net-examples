// Title: Move a sparkline to a new cell range in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that relocates an existing sparkline from cell E1 to cell G1 using Aspose.Cells SparklineGroup.ResetRanges. | Show how to programmatically change the destination range of a line sparkline while keeping its source data intact in Aspose.Cells. | Demonstrate moving a sparkline group to a different cell area without recreating the workbook.
// Common Searches: Aspose.Cells C# how to change sparkline destination cell | reset sparkline ranges to move sparkline in Excel using .NET SDK | move line sparkline from column E to column G programmatically | C# example for relocating sparkline group with Aspose.Cells | update sparkline location after adding data in Aspose.Cells workbook
// Tags: sparkline relocation using Aspose.Cells ResetRanges | C# sparkline group reposition to new cell area | Aspose.Cells change sparkline destination range | programmatic sparkline repositioning Excel .NET | reset sparkline ranges example C#

using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds sample data, inserts a line sparkline in cell E1 based on A1:D1, then uses SparklineGroup.ResetRanges to move the sparkline to cell G1, and finally saves the file as MovedSparkline.xlsx.
class MoveSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the original location of the sparkline (cell E1)
        CellArea originalLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group with the original location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Define the new location for the sparkline (cell G1)
        CellArea newLocation = CellArea.CreateCellArea(0, 6, 0, 6); // Column G (0‑based index)

        // Move the sparkline by resetting its ranges.
        // This clears the existing sparkline and creates a new one at the specified location.
        group.ResetRanges("A1:D1", false, newLocation);

        // Save the workbook with the moved sparkline
        workbook.Save("MovedSparkline.xlsx");
    }
}
