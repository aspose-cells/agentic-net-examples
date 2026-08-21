// Title: Aspose.Cells for .NET – Set Red High‑Point Marker on a Line Sparkline in Cell H3 (C#)
// Description: C# example that creates a workbook, adds a line sparkline for range A1:D1, places it in H3, enables markers, highlights the high point and colors that marker red, then saves the file as SparklineHighPointMarker_H3.xlsx.
// Keywords: Aspose.Cells sparkline high point color | C# line sparkline red marker | set sparkline marker color Aspose.Cells | .NET sparkline custom marker | sparkline high point highlight
// Common Searches: Aspose.Cells change high‑point marker color C# | add line sparkline to cell H3 Aspose.Cells | show high point in sparkline .NET | customize sparkline markers Aspose.Cells | red high‑point sparkline example
// Developer Intent: Add a line sparkline to H3 and display its highest value with a red marker.
// Use Cases: Financial dashboards that flag peak values in red for instant visual cues. | Automated reporting where each trend line highlights its maximum point. | Data‑quality sheets that use red high‑point markers to identify outliers.
// AI Prompts: Generate C# code using Aspose.Cells to insert a line sparkline in cell H3, turn on markers, show the high point, and set the high‑point marker color to red. | Explain how to customize sparkline marker colors (high‑point, low‑point, negative) with Aspose.Cells for .NET. | Provide step‑by‑step instructions for creating a sparkline group, adding a sparkline, enabling markers, and applying a red high‑point marker in a workbook.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, adds a line sparkline for range A1:D1, places it in H3, enables markers, highlights the high point and colors that marker red, then saves the file as SparklineHighPointMarker_H3.xlsx.
class SparklineHighPointMarkerDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (A1:D1)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(3);

        // Define the location of the sparkline: cell H3 (column 7, row 2)
        CellArea sparklineLocation = new CellArea
        {
            StartColumn = 7, // H
            EndColumn = 7,
            StartRow = 2,    // 3rd row (zero‑based)
            EndRow = 2
        };

        // Add a line sparkline group using the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (the same data range)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 7);

        // Enable markers and high‑point highlighting
        group.ShowMarkers = true;
        group.ShowHighPoint = true;

        // Set the high‑point marker color to red
        CellsColor highPointColor = workbook.CreateCellsColor();
        highPointColor.Color = Color.Red;
        group.HighPointColor = highPointColor;

        // Save the workbook
        workbook.Save("SparklineHighPointMarker_H3.xlsx", SaveFormat.Xlsx);
    }
}
