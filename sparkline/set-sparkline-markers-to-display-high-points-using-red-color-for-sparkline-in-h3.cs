// Title: How to set red high‑point markers for a line sparkline placed in cell H3 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set high point marker color for a sparkline in a specific cell | change sparkline high point marker to red with Aspose.Cells .NET | add line sparkline to H3 and highlight highest value with red marker using Aspose.Cells | C# Aspose.Cells sparkline group show high point and apply custom color | customize sparkline high point color in an Excel workbook via Aspose.Cells API
// Tags: Aspose.Cells sparkline high point color | line sparkline marker customization .NET | set sparkline high point color C# | add sparkline to cell H3 Aspose.Cells | Excel sparkline high point red Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, writes values to A3:D3, inserts a line sparkline into cell H3, enables the high‑point marker, sets its color to red using a CellsColor object, and saves the file as SparklineHighPointRed.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data in row 3 (A3:D3)
        sheet.Cells["A3"].PutValue(5);
        sheet.Cells["B3"].PutValue(12);
        sheet.Cells["C3"].PutValue(8);
        sheet.Cells["D3"].PutValue(15);

        // Define the cell area where the sparkline will be placed (H3)
        CellArea sparklineArea = new CellArea
        {
            StartRow = 2,   // Row index is zero‑based (2 => row 3)
            EndRow = 2,
            StartColumn = 7, // Column H is index 7
            EndColumn = 7
        };

        // Add a line sparkline group that uses the data range A3:D3
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,          // Sparkline type
            "A3:D3",                     // Data range
            false,                       // Plot right‑to‑left
            sparklineArea);              // Destination cell area

        // Retrieve the created sparkline group
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Enable high‑point markers
        group.ShowHighPoint = true;

        // Set the high‑point marker color to red
        CellsColor redColor = workbook.CreateCellsColor();
        redColor.Color = Color.Red;
        group.HighPointColor = redColor;

        // (Optional) Show low point as well for contrast
        // group.ShowLowPoint = true;
        // group.LowPointColor = workbook.CreateCellsColor();
        // group.LowPointColor.Color = Color.Blue;

        // Save the workbook
        string outputPath = "SparklineHighPointRed.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
