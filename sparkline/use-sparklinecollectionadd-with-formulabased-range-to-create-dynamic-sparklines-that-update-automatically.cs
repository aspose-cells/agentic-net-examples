// Title: Add line sparklines with OFFSET formula ranges using SparklineCollection.Add in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add sparkline with OFFSET formula for each row | create dynamic sparklines that update when source data changes in .NET | use SparklineCollection.Add to set formula‑based data range in Excel workbook | programmatically generate line sparklines with formula ranges using Aspose.Cells | refresh sparkline data after modifying worksheet cells in C#
// Tags: SparklineCollection.Add with OFFSET formula | dynamic line sparklines Aspose.Cells | formula‑based sparkline data range .NET | auto‑updating sparklines Excel C# | SparklineGroup programmatic source range

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace DynamicSparklineDemo
{
    // The example creates a new workbook, fills a 5 × 4 data table, adds a line sparkline group in column E, and inserts a sparkline for each row using an OFFSET formula so the source range updates automatically. After changing the values of the second row, it recalculates formulas to refresh the sparklines, prints each sparkline's data range, and saves the workbook as DynamicSparkline.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // 2. Populate sample data (5 rows, 4 columns)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                    }
                }

                // 3. Define the location area where sparklines will be placed (E1:E5)
                CellArea locationArea = CellArea.CreateCellArea("E1", "E5");

                // 4. Add a SparklineGroup (initial range will be overridden per sparkline)
                int groupIdx = sheet.SparklineGroups.Add(
                    SparklineType.Line,
                    "A1:D1",          // dummy initial range
                    false,            // plot empty cells as gaps
                    locationArea);    // where sparklines appear

                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIdx];
                SparklineCollection sparklineCollection = sparklineGroup.Sparklines;

                // 5. Add a sparkline for each data row using an OFFSET formula
                for (int i = 0; i < 5; i++)
                {
                    // OFFSET(DataSheet!$A$1, rowOffset, 0, 1, 4)
                    string formulaRange = $"=OFFSET(DataSheet!$A$1,{i},0,1,4)";
                    sparklineCollection.Add(formulaRange, i, 4); // column 4 = E
                }

                // 6. Modify source data for row 2 (zero‑based index 1)
                Console.WriteLine("Modifying source data for row 2 (zero‑based index 1)...");
                sheet.Cells[1, 0].PutValue(100); // A2
                sheet.Cells[1, 1].PutValue(200); // B2
                sheet.Cells[1, 2].PutValue(300); // C2
                sheet.Cells[1, 3].PutValue(400); // D2

                // Recalculate formulas so that OFFSET ranges are refreshed
                workbook.CalculateFormula();

                // 7. Output each sparkline's data range to verify formulas
                Console.WriteLine("Sparkline data ranges after modification:");
                for (int i = 0; i < sparklineCollection.Count; i++)
                {
                    Sparkline sp = sparklineCollection[i];
                    Console.WriteLine($"Row {i + 1} (Cell {sp.Row + 1},{sp.Column + 1}): {sp.DataRange}");
                }

                // 8. Save the workbook
                string outFile = "DynamicSparkline.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outFile}'. Open it in Excel to see dynamic sparklines.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
