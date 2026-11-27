// Title: Setting the 3‑D depth of a column sparkline with Aspose.Cells for .NET (when the API supports it)
// Common Searches: Aspose.Cells C# set column sparkline 3d depth | increase visual depth of Excel sparkline using Aspose.Cells .NET | how to adjust sparkline depth property in a workbook with Aspose.Cells | C# example for configuring 3d depth of sparkline group in Aspose.Cells | Aspose.Cells sparkline depth not available in current version workaround
// Tags: Aspose.Cells sparkline depth property | C# column sparkline 3d configuration | Excel sparkline visual customization Aspose.Cells | Aspose.Cells workbook sparkline group styling | 3d sparkline depth adjustment .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDepthDemo
{
    // The example creates a workbook, fills a 4×5 data range, and adds a column sparkline group (which can render in 3‑D) to cells G1‑G4. It shows where the Depth property would be set if supported by the Aspose.Cells version, and demonstrates additional styling such as series color, high‑point and low‑point colors, and line weight before saving the file as XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            string outputPath = "SparklineDepthDemo.xlsx";

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (4 rows, 5 columns)
                for (int row = 0; row < 4; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                    }
                }

                // Define the area where the sparkline will be displayed (column G, rows 1‑4)
                CellArea sparklineArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = 3,
                    StartColumn = 6,
                    EndColumn = 6
                };

                // Add a column sparkline group (3‑D column sparkline)
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Column,   // Column sparkline (supports 3‑D depth)
                    "A1:E4",                // Data range
                    false,                  // Do not plot right‑to‑left
                    sparklineArea           // Destination area
                );

                // Retrieve the created SparklineGroup
                SparklineGroup sparkGroup = sheet.SparklineGroups[groupIndex];

                // NOTE: The Depth property is not available in the current Aspose.Cells version.
                // If a newer version provides it, you can set it here, e.g., sparkGroup.Depth = 10;

                // Customize visual aspects
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.LightSeaGreen;
                sparkGroup.SeriesColor = seriesColor;

                sparkGroup.ShowHighPoint = true;
                CellsColor highColor = workbook.CreateCellsColor();
                highColor.Color = Color.Gold;
                sparkGroup.HighPointColor = highColor;

                sparkGroup.ShowLowPoint = true;
                CellsColor lowColor = workbook.CreateCellsColor();
                lowColor.Color = Color.Crimson;
                sparkGroup.LowPointColor = lowColor;

                sparkGroup.LineWeight = 1.5; // Slightly thicker lines

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
