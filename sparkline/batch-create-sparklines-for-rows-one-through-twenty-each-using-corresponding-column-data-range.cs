// Title: Create line sparklines for rows 1‑20 from each column’s data range and place them in column V using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a line sparkline for each row based on its column data | how to generate multiple sparklines in a worksheet with Aspose.Cells .NET | Aspose.Cells create sparkline groups for rows 1 to 20 in column V | batch create sparklines from vertical ranges using Aspose.Cells C# example | customize sparkline colors and markers in Aspose.Cells C# code
// Tags: create line sparkline groups Aspose.Cells | sparkline batch generation C# | vertical data range sparkline Aspose.Cells | place sparklines in column V Aspose.Cells | customize sparkline colors Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineBatchDemo
{
    // The sample program builds a new workbook, fills cells A1:T20 with random numbers, then loops through rows 1‑20 to add a line sparkline group for each row. Each sparkline uses the vertical range of its corresponding column (e.g., A1:A20) and is placed in column V. The sparkline groups are customized to show high and low points with green and red markers, a blue series line, and a line weight of 0.75. Finally, the workbook is saved as SparklinesBatch.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in columns A to T (20 columns) and rows 1 to 20
            Random rnd = new Random();
            for (int row = 0; row < 20; row++)          // rows 0‑19 correspond to Excel rows 1‑20
            {
                for (int col = 0; col < 20; col++)      // columns 0‑19 correspond to A‑T
                {
                    cells[row, col].PutValue(rnd.Next(1, 100));
                }
            }

            // Column index where sparklines will be placed (V = 22nd column, zero‑based index 21)
            const int sparklineColumnIndex = 21; // V

            // Loop through rows 1‑20 and create a sparkline for each row using its own column data range
            for (int i = 0; i < 20; i++) // i = 0‑19 => Excel rows 1‑20
            {
                // Determine the column letter for the current row (A‑T)
                string dataColumnLetter = CellsHelper.ColumnIndexToName(i);
                // Build the vertical data range for this column (e.g., "A1:A20")
                string dataRange = $"{dataColumnLetter}1:{dataColumnLetter}20";

                // Define the location cell for the sparkline (V row i+1)
                CellArea location = new CellArea
                {
                    StartRow = i,
                    EndRow = i,
                    StartColumn = sparklineColumnIndex,
                    EndColumn = sparklineColumnIndex
                };

                // Add a sparkline group: Line type, vertical data, placed at the location cell
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, true, location);

                // Optional: customize the sparkline group (e.g., color, markers)
                SparklineGroup group = sheet.SparklineGroups[groupIndex];
                group.ShowHighPoint = true;
                group.ShowLowPoint = true;
                group.HighPointColor.Color = System.Drawing.Color.Green;
                group.LowPointColor.Color = System.Drawing.Color.Red;
                group.SeriesColor.Color = System.Drawing.Color.Blue;
                group.LineWeight = 0.75;
            }

            // Save the workbook
            string outputPath = "SparklinesBatch.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
