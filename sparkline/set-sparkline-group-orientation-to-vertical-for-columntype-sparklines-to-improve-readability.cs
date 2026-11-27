// Title: Configure a column sparkline group to display vertically in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column sparkline orientation to vertical | vertical column sparklines with Aspose.Cells .NET example | change PlotRightToLeft property for sparkline group in C# | create vertical column sparklines in Excel using Aspose.Cells | customize high and low point colors for vertical sparklines Aspose.Cells
// Tags: column sparkline orientation flag | vertical sparkline group setup | Aspose.Cells sparkline high low point styling | export Excel workbook containing vertical sparklines | C# Aspose.Cells sparkline group API

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample creates a workbook, fills a 4x5 data range, adds a column sparkline group linked to that range, sets the PlotRightToLeft flag to false to make the sparklines vertical, customizes high/low point colors and series color, and saves the file as SparklineVerticalColumn.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (4 rows, 5 columns)
        // Data will be placed in columns A to E (0‑4) rows 0‑3
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // Define the area where the sparklines will be placed.
        // We'll put a column sparkline for each row in column G (index 6)
        CellArea sparklineArea = new CellArea
        {
            StartRow = 0,
            EndRow = 3,
            StartColumn = 6,
            EndColumn = 6
        };

        // Add a sparkline group of type Column.
        // Data range: A1:E4 (columns 0‑4, rows 0‑3)
        // The third argument (false) means the data is plotted left‑to‑right.
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Column,
            "A1:E4",
            false,
            sparklineArea);

        // Retrieve the created group
        SparklineGroup columnGroup = sheet.SparklineGroups[groupIndex];

        // Set the orientation to vertical.
        // In Aspose.Cells this is controlled by the PlotRightToLeft property.
        // False = vertical (default), True = horizontal.
        columnGroup.PlotRightToLeft = false;

        // Optional: customize appearance
        columnGroup.ShowHighPoint = true;
        columnGroup.ShowLowPoint = true;
        columnGroup.HighPointColor.Color = System.Drawing.Color.Green;
        columnGroup.LowPointColor.Color = System.Drawing.Color.Red;
        columnGroup.SeriesColor = workbook.CreateCellsColor();
        columnGroup.SeriesColor.Color = System.Drawing.Color.DarkBlue;

        // Save the workbook
        string outputPath = "SparklineVerticalColumn.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to '{outputPath}'. Sparkline group orientation set to vertical.");
    }
}
