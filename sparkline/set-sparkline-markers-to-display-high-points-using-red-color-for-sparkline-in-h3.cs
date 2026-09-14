// Title: How to add a line sparkline to cell H3 and set the high‑point marker color to red using Aspose.Cells for .NET (C#)
// AI Prompts: Create a line sparkline in cell H3 that references range A3:D3, enable markers, highlight the high point, and set its marker color to red with Aspose.Cells in C#. | Generate sample data in row 3, add a sparkline group, configure ShowHighPoint and HighPointColor properties, and save the workbook as an XLSX file using Aspose.Cells. | Write C# code that applies a red high‑point marker to a line sparkline placed in H3 of a worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells C# set red color for high point marker in a line sparkline | How to place a sparkline in cell H3 and highlight the highest value using Aspose.Cells | C# example for adding a line sparkline with markers and custom high‑point color | Saving a workbook with a sparkline that has red high‑point markers in Aspose.Cells .NET
// Tags: Aspose.Cells line sparkline high‑point color | C# set sparkline high point marker red | Aspose.Cells add sparkline to specific cell | C# configure sparkline markers in XLSX workbook | Aspose.Cells sparkline group properties

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells A3:D3 with data, adds a line sparkline to cell H3, enables markers, highlights the highest point, sets the high‑point marker color to red, and saves the file as SparklineHighPointMarker.xlsx.
class SparklineHighPointMarkerDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 3, columns A to D)
        sheet.Cells["A3"].PutValue(5);
        sheet.Cells["B3"].PutValue(2);
        sheet.Cells["C3"].PutValue(8);
        sheet.Cells["D3"].PutValue(3);

        // Define the location of the sparkline (cell H3)
        CellArea sparklineLocation = new CellArea
        {
            StartColumn = 7, // Column H (0‑based index)
            EndColumn = 7,
            StartRow = 2,    // Row 3 (0‑based index)
            EndRow = 2
        };

        // Add a line sparkline group that uses the data range A3:D3 and places the sparkline in H3
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A3:D3", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (required for the group to contain an item)
        group.Sparklines.Add(sheet.Name + "!A3:D3", 2, 7);

        // Enable markers and highlight the highest points
        group.ShowMarkers = true;
        group.ShowHighPoint = true;

        // Set the high‑point marker color to red
        CellsColor redColor = workbook.CreateCellsColor();
        redColor.Color = Color.Red;
        group.HighPointColor = redColor;

        // Save the workbook
        workbook.Save("SparklineHighPointMarker.xlsx", SaveFormat.Xlsx);
    }
}
