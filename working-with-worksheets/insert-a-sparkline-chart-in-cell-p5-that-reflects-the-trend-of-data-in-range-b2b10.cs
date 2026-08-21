// Title: C# – Insert a Line Sparkline in Cell P5 from Range B2:B10 using Aspose.Cells
// Description: Creates a new workbook, adds a line‑type SparklineGroup to the first worksheet, places a sparkline that references B2:B10 into cell P5, and saves the file as SparklineInCell.xlsx.
// Keywords: Aspose.Cells C# sparkline | line sparkline cell P5 | sparkline range B2:B10 | add sparkline programmatically | Excel sparkline Aspose
// Common Searches: how to add a sparkline to a single cell with Aspose.Cells | C# sparkline from B2 to B10 placed in P5 | Aspose.Cells create line sparkline in Excel | insert sparkline chart into worksheet cell using .NET
// Developer Intent: Add a line sparkline to cell P5 that visualizes the values in B2:B10.
// Use Cases: Show a compact trend line next to a data table in a dashboard. | Automate sparklines for many rows by iterating over cell coordinates. | Include visual trend indicators in generated Excel reports.
// AI Prompts: Generate code to add multiple sparklines for different ranges in adjacent cells with Aspose.Cells. | Show how to customize sparkline color, weight, and markers after creation. | Explain how to retrieve an existing SparklineGroup and change its source data range.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds a line‑type SparklineGroup to the first worksheet, places a sparkline that references B2:B10 into cell P5, and saves the file as SparklineInCell.xlsx.
class SparklineInCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a sparkline group of type Line
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

        // Add a sparkline that uses the data range B2:B10
        // Place the sparkline in cell P5 (zero‑based row 4, column 15)
        sparklineGroup.Sparklines.Add("B2:B10", 4, 15);

        // Save the workbook
        workbook.Save("SparklineInCell.xlsx");
    }
}
