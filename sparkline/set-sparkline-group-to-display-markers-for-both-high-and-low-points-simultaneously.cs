// Title: Enable high‑ and low‑point markers in a line sparkline group using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# show high point marker in line sparkline | how to display low point marker for sparkline group with Aspose.Cells | set both high and low markers in Excel sparkline using Aspose.Cells .NET | customize sparkline high and low marker colors in C# Aspose.Cells | add line sparkline group A1:E5 column F with markers Aspose.Cells
// Tags: Aspose.Cells line sparkline high marker | Aspose.Cells sparkline low marker | sparkline group marker color settings Aspose.Cells | C# ShowHighPoint ShowLowPoint Aspose.Cells | Excel sparkline markers setup .NET | Aspose.Cells add sparkline group with markers

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace SparklineHighLowMarkersDemo
{
    // The example creates a workbook, fills a 5x5 range, adds a line sparkline group covering A1:E5 placed in column F, enables both high‑ and low‑point markers, assigns custom colors to each marker, adjusts line weight and series color, and saves the file as SparklineHighLowMarkers.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (5 rows x 5 columns)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        // Example values: (row+1)*(col+1) with some variation
                        sheet.Cells[row, col].PutValue((row + 1) * (col + 1) + (row == col ? 5 : 0));
                    }
                }

                // Define the area where the sparklines will be placed (column F, rows 1-5)
                CellArea sparklineArea = new CellArea
                {
                    StartRow = 0,   // zero‑based index (row 1)
                    EndRow = 4,     // row 5
                    StartColumn = 5, // column F (0‑based)
                    EndColumn = 5
                };

                // Add a line sparkline group that uses the data range A1:E5
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Line,   // type of sparkline
                    "A1:E5",              // data range
                    false,                // plot right‑to‑left = false
                    sparklineArea);       // where the sparklines will appear

                // Retrieve the created sparkline group
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Enable markers for both high and low points
                group.ShowHighPoint = true;
                group.ShowLowPoint = true;

                // Optional: set distinct colors for the high and low markers
                group.HighPointColor.Color = Color.Green;
                group.LowPointColor.Color = Color.Red;

                // Optional: customize other appearance aspects
                group.LineWeight = 1.0;               // line thickness
                group.SeriesColor.Color = Color.Blue; // line color

                // Save the workbook
                string outputPath = "SparklineHighLowMarkers.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Sparkline group displays both high and low point markers.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the sparkline workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
