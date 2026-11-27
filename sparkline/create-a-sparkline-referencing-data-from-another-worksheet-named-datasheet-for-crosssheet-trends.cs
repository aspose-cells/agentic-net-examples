// Title: Generate a line sparkline on a summary worksheet that pulls data from a separate 'DataSheet' sheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create sparkline that references another worksheet | How to add a cross-sheet line sparkline in a .NET workbook with Aspose.Cells | C# example for SparklineGroup using DataSheet!A2:A6 range in Aspose.Cells | Set high and low point colors for a sparkline in Aspose.Cells C# | Save workbook with sparkline on a summary sheet using Aspose.Cells .NET
// Tags: create cross-sheet sparkline Aspose.Cells | line sparkline group C# | sparkline data range with sheet name Aspose.Cells | customize sparkline high low points C# | save workbook with sparkline Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CrossSheetSparklineDemo
{
    // The sample creates a new workbook, adds a worksheet named 'DataSheet' and fills cells A1:A5 with sample values. It then adds a 'Summary' worksheet, defines a sparkline area in column F, and creates a line SparklineGroup that references the range 'DataSheet!A2:A6'. Individual sparklines are added for rows 2‑6, high and low point markers are enabled with green and red colors, line weight is set, and the workbook is saved as 'CrossSheetSparkline.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -----------------------------------------------------------------
                // 1. Create and populate the source worksheet named "DataSheet"
                // -----------------------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "DataSheet";

                // Fill some sample data (A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    // Example values: 10, 20, 15, 30, 25
                    dataSheet.Cells[i, 0].PutValue((i + 1) * 10);
                }

                // -----------------------------------------------------------------
                // 2. Add a second worksheet that will contain the sparkline
                // -----------------------------------------------------------------
                int summaryIndex = workbook.Worksheets.Add();
                Worksheet summarySheet = workbook.Worksheets[summaryIndex];
                summarySheet.Name = "Summary";

                // Optional: label the sparkline column
                summarySheet.Cells["F1"].PutValue("Trend");

                // -----------------------------------------------------------------
                // 3. Define the area where the sparkline(s) will be placed
                //    We'll place them in column F (index 5), rows 1‑5 (0‑4 zero‑based)
                // -----------------------------------------------------------------
                CellArea sparklineArea = new CellArea
                {
                    StartRow = 1,      // Row 2 in Excel (zero‑based index)
                    EndRow = 5,        // Row 6 in Excel
                    StartColumn = 5,   // Column F
                    EndColumn = 5
                };

                // -----------------------------------------------------------------
                // 4. Add a sparkline group that references the data on "DataSheet"
                //    Data range string must include the sheet name, e.g. "DataSheet!A2:A6"
                // -----------------------------------------------------------------
                string dataRange = "DataSheet!A2:A6"; // Corresponds to rows 2‑6 (A2:A6)
                int groupIndex = summarySheet.SparklineGroups.Add(
                    SparklineType.Line,   // Type of sparkline
                    dataRange,            // Cross‑sheet data range
                    false,                // isVertical = false (default)
                    sparklineArea);       // Where the sparklines will appear

                SparklineGroup sparklineGroup = summarySheet.SparklineGroups[groupIndex];

                // -----------------------------------------------------------------
                // 5. Add individual sparklines to the group (one per row)
                //    The Add method requires the same data range and the target cell.
                // -----------------------------------------------------------------
                for (int row = 1; row <= 5; row++)
                {
                    // Each sparkline will be placed in column F of its respective row
                    sparklineGroup.Sparklines.Add(dataRange, row, 5);
                }

                // -----------------------------------------------------------------
                // 6. Optional: customize appearance
                // -----------------------------------------------------------------
                sparklineGroup.ShowHighPoint = true;
                sparklineGroup.ShowLowPoint = true;
                sparklineGroup.HighPointColor.Color = System.Drawing.Color.Green;
                sparklineGroup.LowPointColor.Color = System.Drawing.Color.Red;
                sparklineGroup.LineWeight = 1.0;

                // -----------------------------------------------------------------
                // 7. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "CrossSheetSparkline.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the sparkline:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
