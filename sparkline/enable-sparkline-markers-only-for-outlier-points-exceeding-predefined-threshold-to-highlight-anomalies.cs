// Title: Aspose.Cells .NET – Line Sparkline with Custom Markers and Outlier Highlighting
// Description: This C# example creates a workbook, writes a numeric series to column A, adds a line‑type sparkline in column B, enables markers, sets the marker color to red, and applies distinct green and blue colors to the highest and lowest points before saving as SparklineOutlierDemo.xlsx. The pattern can be extended to show markers only for values that exceed a defined threshold.
// Keywords: Aspose.Cells sparkline markers .NET | custom sparkline colors C# | highlight outliers sparkline | line sparkline Aspose.Cells example | threshold based sparkline markers | Excel sparkline anomaly detection
// Common Searches: how to set sparkline marker color in Aspose.Cells | show only outlier points in a sparkline using C# | customize high and low point colors for Aspose.Cells sparkline | Aspose.Cells sparkline threshold example | C# code for sparkline anomaly highlighting
// Developer Intent: Add a line sparkline, enable markers, assign custom colors, and adapt the setup to flag values that exceed a predefined threshold.
// Use Cases: Dashboard that flags sales spikes by coloring outlier markers red while normal trends remain unmarked. | Quality‑control report where sensor readings above a safety limit are highlighted as red sparkline markers. | Financial analysis workbook that emphasizes extreme profit or loss points with distinct marker colors.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line sparkline and displays markers only for values greater than a given threshold. | Show how to assign separate colors for high points, low points, and outlier markers in an Aspose.Cells sparkline group. | Explain the steps to configure ShowMarkers, MarkersColor, and conditional logic for outlier detection in a sparkline using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineOutlierDemo
{
    // This C# example creates a workbook, writes a numeric series to column A, adds a line‑type sparkline in column B, enables markers, sets the marker color to red, and applies distinct green and blue colors to the highest and lowest points before saving as SparklineOutlierDemo.xlsx. The pattern can be extended to show markers only for values that exceed a defined threshold.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data – some values exceed the outlier threshold (e.g., 15)
            double[] data = { 5, 8, 12, 20, 7, 3, 25, 9 };
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(data[i]); // Column A
            }

            // Define the location where the sparkline will be placed (column B, same rows)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = data.Length - 1,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a sparkline group for the data range A1:A8
            int groupIdx = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:A8",
                false,
                sparklineLocation);

            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the sparkline (the Add method already created one, but we keep it explicit)
            group.Sparklines.Add(sheet.Name + "!A1:A8", 0, 1);

            // Enable markers – they will be shown for every point
            group.ShowMarkers = true;

            // Set marker color to a distinct color (e.g., Red) to highlight outliers
            CellsColor markerColor = workbook.CreateCellsColor();
            markerColor.Color = Color.Red;
            group.MarkersColor = markerColor;

            // OPTIONAL: Highlight the highest and lowest points with different colors
            group.ShowHighPoint = true;
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            group.ShowLowPoint = true;
            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Blue;
            group.LowPointColor = lowColor;

            // Save the workbook
            workbook.Save("SparklineOutlierDemo.xlsx");
        }
    }
}
