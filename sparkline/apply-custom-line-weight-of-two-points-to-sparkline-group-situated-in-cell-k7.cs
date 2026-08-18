// Title: C# – Set Sparkline LineWeight to 2 pts in cell K7 with Aspose.Cells
// Description: Shows how to create a workbook, add a line sparkline for range A7:D7, place it in cell K7, and set SparklineGroup.LineWeight to 2 points using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# sparkline line weight | SparklineGroup.LineWeight | line sparkline thickness | set sparkline line weight | cell K7 sparkline | Aspose.Cells example | Excel sparkline styling .NET | custom sparkline line thickness
// Common Searches: Aspose.Cells set sparkline line weight C# | How to change sparkline thickness in .NET | LineWeight property SparklineGroup example | Place sparkline in specific cell Aspose.Cells | Increase sparkline line thickness to 2 points
// Developer Intent: The developer wants to programmatically set the line thickness of a line sparkline group to 2 points and locate the sparkline in cell K7 using Aspose.Cells for .NET.
// Use Cases: Financial dashboards where a uniform 2‑point sparkline line improves readability. | Automated reporting pipelines that enforce brand‑specific sparkline thickness across multiple worksheets. | User‑customizable Excel exports that adjust sparkline appearance based on runtime parameters.
// AI Prompts: Generate C# code that loads an existing workbook and updates all line sparklines to a specified LineWeight value. | Explain the valid range for SparklineGroup.LineWeight and how different values affect rendering in Excel. | Provide an example that applies conditional line weight to sparklines based on data thresholds using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add a line sparkline for range A7:D7, place it in cell K7, and set SparklineGroup.LineWeight to 2 points using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (cells A7:D7)
        sheet.Cells["A7"].PutValue(5);
        sheet.Cells["B7"].PutValue(2);
        sheet.Cells["C7"].PutValue(1);
        sheet.Cells["D7"].PutValue(3);

        // Define the location where the sparkline will be placed (cell K7)
        CellArea location = new CellArea
        {
            StartColumn = 10, // Column K (0‑based index)
            EndColumn   = 10,
            StartRow    = 6,  // Row 7 (0‑based index)
            EndRow      = 6
        };

        // Add a line sparkline group with the data range A7:D7 and place it at K7
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A7:D7", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Apply a custom line weight of 2 points to the sparkline group
        group.LineWeight = 2.0;

        // Save the workbook
        workbook.Save("SparklineLineWeightK7.xlsx", SaveFormat.Xlsx);
    }
}
