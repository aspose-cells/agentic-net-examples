using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCrossSheet
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet that will hold the source data
                Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                dataSheet.Name = "DataSheet";

                // Populate sample data in DataSheet (A1:D5)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        dataSheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                    }
                }

                // Add a worksheet where the sparkline will be placed
                Worksheet trendSheet = workbook.Worksheets[0]; // default first sheet
                trendSheet.Name = "TrendSheet";

                // Define the location range for the sparkline group (one sparkline per column of data)
                // Since the data range has 4 columns, we need 4 cells to host the sparklines (E1:H1)
                CellArea sparklineLocation = CellArea.CreateCellArea("E1", "H1");

                // Data range format: SheetName!StartCell:EndCell
                string dataRange = "DataSheet!A1:D5";

                // Add a sparkline group on TrendSheet that references the cross‑sheet data range
                int groupIndex = trendSheet.SparklineGroups.Add(
                    SparklineType.Line,   // sparkline type
                    dataRange,            // cross‑sheet data range
                    false,                // plot by column (false = by column)
                    sparklineLocation);   // where the sparklines will appear

                // Retrieve the created group (optional: customize appearance)
                SparklineGroup group = trendSheet.SparklineGroups[groupIndex];
                group.ShowHighPoint = true;
                group.ShowLowPoint = true;

                // Define output file path
                string outputPath = "CrossSheetSparkline.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}