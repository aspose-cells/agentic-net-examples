// Title: Remove the second sparkline from a SparklineGroup with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, fills three rows of numeric data, adds a line‑type SparklineGroup covering A1:D3 with sparklines placed in column E, then deletes the sparkline at index 1 using SparklineGroup.Sparklines.RemoveAt, and finally saves the file as DeleteSecondSparkline.xlsx.
// Keywords: Aspose.Cells | C# sparkline example | SparklineGroup RemoveAt | delete specific sparkline | Excel sparkline manipulation | Aspose.Cells tutorial | programmatic sparkline removal | line sparkline group | Excel automation C# | Aspose.Cells API
// Common Searches: Aspose.Cells remove sparkline by index | C# SparklineGroup RemoveAt usage | how to delete a specific sparkline in Excel with Aspose | remove second sparkline from SparklineGroup | Aspose.Cells SparklineGroup example C#
// Developer Intent: Delete the sparkline at position 1 in the first SparklineGroup.
// Use Cases: Clean up automatically generated sparklines that are no longer needed before exporting a report. | Adjust the number of displayed sparklines based on user‑selected data ranges in a dashboard. | Programmatically eliminate invalid or outdated sparklines when synchronizing workbook data.
// AI Prompts: Generate C# code that removes the sparkline at index 1 from a SparklineGroup using Aspose.Cells. | Explain how SparklineGroup.Sparklines.RemoveAt reindexes the remaining sparklines. | Show an example that loops through a SparklineGroup and deletes sparklines that meet a custom condition.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, fills three rows of numeric data, adds a line‑type SparklineGroup covering A1:D3 with sparklines placed in column E, then deletes the sparkline at index 1 using SparklineGroup.Sparklines.RemoveAt, and finally saves the file as DeleteSecondSparkline.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (3 rows, 4 columns) to generate multiple sparklines
        for (int row = 0; row < 3; row++)
        {
            worksheet.Cells[row, 0].PutValue(row + 1);
            worksheet.Cells[row, 1].PutValue((row + 1) * 2);
            worksheet.Cells[row, 2].PutValue((row + 1) * 3);
            worksheet.Cells[row, 3].PutValue((row + 1) * 4);
        }

        // Define the location range where sparklines will be placed (column E, rows 0‑2)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 2,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group; this creates one sparkline per row in the data range
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D3", false, location);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // The group now contains three sparklines (indices 0, 1, 2).
        // Delete the second sparkline (index 1) using RemoveAt.
        sparklineGroup.Sparklines.RemoveAt(1);

        // Save the workbook to verify the removal.
        workbook.Save("DeleteSecondSparkline.xlsx", SaveFormat.Xlsx);
    }
}
