using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 1. Prepare the source data on a worksheet named "DataSheet"
            // -----------------------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "DataSheet";

            // Fill sample data (5 rows × 4 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    dataSheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // -----------------------------------------------------------------
            // 2. Add a worksheet that will host the sparkline
            // -----------------------------------------------------------------
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Define the cells where the sparkline(s) will appear (one per row)
            CellArea sparklineLocation = CellArea.CreateCellArea("F1", "F5");

            // -----------------------------------------------------------------
            // 3. Create a sparkline group that references the data on "DataSheet"
            // -----------------------------------------------------------------
            // Data range must include the sheet name (quoted if needed)
            string dataRange = "'DataSheet'!A1:D5";

            // Add the sparkline group (Line type, horizontal orientation)
            int groupIdx = summarySheet.SparklineGroups.Add(
                SparklineType.Line,   // sparkline type
                dataRange,            // cross‑sheet data range
                false,                // plot by row (horizontal)
                sparklineLocation);   // where the sparklines will be placed

            // Optional: customize the appearance of the sparkline group
            SparklineGroup group = summarySheet.SparklineGroups[groupIdx];
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            group.SeriesColor = seriesColor;

            // -----------------------------------------------------------------
            // 4. Save the workbook
            // -----------------------------------------------------------------
            string outputPath = "CrossSheetSparkline.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}