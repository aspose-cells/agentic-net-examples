// Title: Create line sparklines in a worksheet and freeze the rows that contain them with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add line sparklines and freeze panes in the same worksheet | how to keep sparklines visible by freezing rows using Aspose.Cells .NET | C# code to insert a sparkline group and apply FreezePanes with Aspose.Cells | freeze rows up to the sparkline range Aspose.Cells example | set high and low point colors for line sparklines in Aspose.Cells C#
// Tags: line sparkline group Aspose.Cells | freeze panes rows Aspose.Cells | customize sparkline high low point colors | populate worksheet random numeric data C# | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace SparklineFreezeDemo
{
    // The sample creates a new workbook, fills cells A1:D10 with random numbers, adds a line sparkline group in column F referencing that data, customizes high/low point colors and line weight, freezes all rows up to the last sparkline using FreezePanes, and saves the file as SparklinesWithFrozenRows.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data (A1:D10) with random numbers
                Random rnd = new Random();
                int dataRows = 10;
                int dataCols = 4; // A, B, C, D
                for (int row = 0; row < dataRows; row++)
                {
                    for (int col = 0; col < dataCols; col++)
                    {
                        sheet.Cells[row, col].PutValue(rnd.Next(1, 100));
                    }
                }

                // 3. Define the source and destination ranges for sparklines
                //    Source: A1:D10, Destination: F1:F10
                string sourceRange = "A1:D10";
                CellArea destination = CellArea.CreateCellArea("F1", "F10");

                // 4. Add a line sparkline group
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Line,   // sparkline type
                    sourceRange,          // source data range
                    false,                // isVertical = false (data series are rows)
                    destination);         // where sparklines will be placed

                SparklineGroup sparkGroup = sheet.SparklineGroups[groupIndex];

                // Optional: customize appearance
                sparkGroup.ShowHighPoint = true;
                sparkGroup.ShowLowPoint = true;
                sparkGroup.HighPointColor.Color = System.Drawing.Color.Green;
                sparkGroup.LowPointColor.Color = System.Drawing.Color.Red;
                sparkGroup.LineWeight = 1.0;
                sparkGroup.SeriesColor = workbook.CreateCellsColor();
                sparkGroup.SeriesColor.Color = System.Drawing.Color.Blue;

                // 5. Freeze rows that contain the sparklines.
                //    Freeze rows up to the row after the last sparkline.
                int freezeRowIndex = destination.EndRow + 1; // zero‑based index of first unfrozen row
                // FreezePanes(row, column, freezedRows, freezedColumns)
                sheet.FreezePanes(freezeRowIndex, 0, freezeRowIndex, 0);

                // 6. Save the workbook
                string outputPath = "SparklinesWithFrozenRows.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Rows 1‑{freezeRowIndex} are frozen to keep the sparklines visible.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
