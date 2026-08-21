// Title: Add a Line Sparkline to Column B from A1:A10 Using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills cells A1‑A10 with values 1‑10, adds a line‑type sparkline group, inserts a sparkline into cell B1 that references A1:A10, and saves the file as SparklineLineExample.xlsx.
// Keywords: Aspose.Cells line sparkline C# | SparklineGroup Sparklines.Add example | add sparkline to column B | Aspose.Cells SparklineType.Line | C# generate sparkline from range
// Common Searches: Aspose.Cells add line sparkline C# | SparklineGroup Sparklines.Add column B | how to create sparkline from A1:A10 using Aspose | C# Aspose.Cells sparkline example
// Developer Intent: Insert a line sparkline into cell B1 that visualizes the data in range A1:A10.
// Use Cases: Show a compact trend line beside raw data in financial reports. | Build dashboards where each row includes a sparkline summarizing its series. | Automate monthly performance sheets that embed line sparklines for quick trend analysis.
// AI Prompts: Generate C# code with Aspose.Cells to add a line sparkline to column C based on range D1:D15 and set the line color to blue. | Explain how to enable markers and adjust line weight for a sparkline group created with Aspose.Cells. | Provide step‑by‑step instructions to add multiple line sparklines to different columns, each referencing its own data range, using SparklineGroup.Sparklines.Add.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, fills cells A1‑A10 with values 1‑10, adds a line‑type sparkline group, inserts a sparkline into cell B1 that references A1:A10, and saves the file as SparklineLineExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:A10
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i + 1); // Column A (index 0)
        }

        // Add a sparkline group of type Line
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Add a sparkline to column B (index 1) at the first row,
        // using the data range A1:A10
        sparklineGroup.Sparklines.Add("A1:A10", 0, 1);

        // Save the workbook
        workbook.Save("SparklineLineExample.xlsx");
    }
}
