using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class MoveRangeWithoutAffectingFrozenPanes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set frozen panes for demonstration (freeze first 3 rows and 3 columns at cell C3)
                worksheet.FreezePanes(3, 3, 3, 3);

                // Create a sample range (A1:B2) and put some data
                AsposeRange sourceRange = worksheet.Cells.CreateRange("A1", "B2");
                sourceRange[0, 0].PutValue("A1");
                sourceRange[0, 1].PutValue("B1");
                sourceRange[1, 0].PutValue("A2");
                sourceRange[1, 1].PutValue("B2");

                // Capture current frozen pane settings
                int frozenRow, frozenColumn, frozenRows, frozenColumns;
                bool hasFreeze = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

                // Unfreeze panes before moving the range to avoid shifting them
                if (hasFreeze)
                {
                    worksheet.UnFreezePanes();
                }

                // Define destination start row and column (move down by 2 rows and right by 2 columns)
                int destRow = sourceRange.FirstRow + 2;
                int destColumn = sourceRange.FirstColumn + 2;

                // Move the range to the destination
                sourceRange.MoveTo(destRow, destColumn);

                // Re‑apply the original frozen pane settings after the move
                if (hasFreeze)
                {
                    worksheet.FreezePanes(frozenRow, frozenColumn, frozenRows, frozenColumns);
                }

                // Save the workbook
                string outputPath = "MoveRangeWithoutAffectingFrozenPanes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            MoveRangeWithoutAffectingFrozenPanes.Run();
        }
    }
}