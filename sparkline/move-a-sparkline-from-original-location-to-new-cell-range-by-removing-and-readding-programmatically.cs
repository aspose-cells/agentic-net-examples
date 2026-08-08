// Title: C# – Relocate a Sparkline to a New Cell with Aspose.Cells for .NET
// Description: Demonstrates how to move a line sparkline from its original cell to another location by removing it from its SparklineGroup and adding it back at the target range, then saving the workbook.
// Keywords: Aspose.Cells sparkline move | C# relocate sparkline | programmatic sparkline reposition .NET | remove and add sparkline Aspose | SparklineGroup manipulation | Excel sparkline location change | MoveSparklineDemo | Aspose.Cells example C#
// Common Searches: how to move a sparkline in Aspose.Cells C# | change sparkline cell address programmatically | Aspose.Cells remove sparkline then add | relocate sparkline to another column .NET | sparkline group reposition example
// Developer Intent: Shift an existing sparkline from its current cell to a different cell by deleting it from the group and inserting it at the new location.
// Use Cases: Re‑arrange dashboard sparklines after inserting or deleting columns. | Align sparklines with newly added summary data in a financial report. | Batch‑move multiple sparklines to a new column range during data model updates.
// AI Prompts: Generate C# code using Aspose.Cells to move a sparkline from E1 to G1 while keeping the same data range. | Explain how to programmatically relocate several sparklines within a SparklineGroup to a new column range in .NET. | Show a step‑by‑step example of removing a sparkline and adding it back at a different cell without recreating the workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to move a line sparkline from its original cell to another location by removing it from its SparklineGroup and adding it back at the target range, then saving the workbook.
class MoveSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the initial location range for the sparkline (cell E1)
        CellArea initialLocation = CellArea.CreateCellArea("E1", "E1");

        // Add a sparkline group with the data range A1:D1 placed at the initial location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, initialLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group at row 0, column 4 (cell E1)
        int sparkIndex = group.Sparklines.Add("A1:D1", 0, 4);
        Sparkline spark = group.Sparklines[sparkIndex];

        // Define the new location for the sparkline (cell G1)
        int newRow = 0;      // same row
        int newColumn = 6;   // column G (0‑based index)

        // Remove the existing sparkline from the group
        group.Sparklines.RemoveSparkline(spark);

        // Re‑add the sparkline at the new location
        int newSparkIndex = group.Sparklines.Add("A1:D1", newRow, newColumn);
        Sparkline newSpark = group.Sparklines[newSparkIndex];

        // Save the workbook with the moved sparkline
        workbook.Save("MovedSparkline.xlsx");
    }
}
