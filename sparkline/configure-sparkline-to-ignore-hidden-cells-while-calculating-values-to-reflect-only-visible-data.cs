// Title: How to create a vertical line sparkline in Aspose.Cells for .NET that skips hidden rows
// AI Prompts: Generate C# code using Aspose.Cells to add a vertical line sparkline that does not count hidden rows in its source range. | Show the steps to set the DisplayHidden flag on a SparklineGroup so hidden cells are excluded from sparkline calculations.
// Common Searches: Aspose.Cells C# sparkline exclude hidden rows example | skip hidden rows when creating sparkline in .NET | DisplayHidden false SparklineGroup Aspose.Cells tutorial | vertical line sparkline from A1:A5 with hidden row 3 ignored
// Tags: Aspose.Cells sparkline ignore hidden rows | SparklineGroup DisplayHidden false | vertical line sparkline Aspose.Cells | C# hide rows sparkline calculation

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills cells A1:A5 with numbers, hides row 3, adds a vertical line sparkline in cell E1 referencing A1:A5, sets the SparklineGroup's DisplayHidden property to false so the hidden row is omitted from calculations, applies a blue series color, and saves the file as SparklineIgnoreHidden.xlsx.
class SparklineIgnoreHiddenDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(3);
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["A5"].PutValue(9);

            // Hide row 3 (index 2) so its value should be ignored by the sparkline
            sheet.Cells.Rows[2].IsHidden = true;

            // Define the cell where the sparkline will be placed (E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group that references the vertical data range A1:A5
            // isVertical = true because the data range is vertical
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", true, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Configure the sparkline to ignore hidden cells when calculating values
            group.DisplayHidden = false; // false = do not include hidden rows/columns

            // Set a series color for visual clarity
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            group.SeriesColor = seriesColor;

            // Save the workbook
            string outputPath = "SparklineIgnoreHidden.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
