// Title: C# – Load Workbook and Insert Line Sparkline Group using Aspose.Cells
// Description: Shows how to open an existing Excel file with Aspose.Cells, select the first worksheet, set a destination range (F1:F4), add a line‑type sparkline group sourced from A1:D4 via the Aspose.Cells.Charts namespace, retrieve the SparklineGroup for optional tweaks, and save the modified workbook as a new file.
// Keywords: Aspose.Cells | C# sparkline example | load workbook Aspose.Cells | add line sparkline | SparklineGroup | Aspose.Cells.Charts | Excel sparkline automation | C# Excel charting | Aspose.Cells .NET | sparkline location range
// Common Searches: how to add a line sparkline to an existing workbook with Aspose.Cells C# | Aspose.Cells load workbook and create sparkline group | using Aspose.Cells.Charts to insert sparklines in C# | save Excel file after adding sparklines Aspose.Cells | C# example for sparkline location range in Aspose.Cells
// Developer Intent: Load an Excel workbook, create a line sparkline group in a specified range, and save the updated file.
// Use Cases: Add trend sparklines to a financial report generated programmatically. | Enhance a dashboard template with visual data cues before distribution. | Automate sparkline insertion for recurring data‑analysis worksheets.
// AI Prompts: Generate C# code that loads a workbook and adds a column sparkline group with custom colors using Aspose.Cells. | Explain how to modify SparklineGroup properties such as markers, weight, and style after creation. | Provide error‑handling examples for invalid source ranges when adding sparklines with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to open an existing Excel file with Aspose.Cells, select the first worksheet, set a destination range (F1:F4), add a line‑type sparkline group sourced from A1:D4 via the Aspose.Cells.Charts namespace, retrieve the SparklineGroup for optional tweaks, and save the modified workbook as a new file.
class Program
{
    static void Main()
    {
        // Path to the existing workbook to be loaded
        string inputPath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet in the workbook
        Worksheet sheet = workbook.Worksheets[0];

        // Define the location range where the sparklines will be placed (e.g., cells F1 to F4)
        CellArea locationRange = CellArea.CreateCellArea("F1", "F4");

        // Add a sparkline group of type Line, using data from A1:D4,
        // not vertical (isVertical = false), and place the sparklines in the defined location range
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D4", false, locationRange);

        // (Optional) Retrieve the created SparklineGroup for further customization
        SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
