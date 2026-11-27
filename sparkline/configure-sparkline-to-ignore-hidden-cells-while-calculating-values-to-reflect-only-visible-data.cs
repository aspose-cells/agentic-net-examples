// Title: How to configure Aspose.Cells sparkline to ignore hidden rows and columns in C#
// Common Searches: Aspose.Cells C# sparkline ignore hidden rows when calculating values | Set DisplayHidden false for sparkline group in Aspose.Cells | Create line sparkline in Aspose.Cells that excludes hidden columns | C# example of hiding rows and columns without affecting sparkline data | Aspose.Cells sparkline group configuration to skip hidden cells | Generate Excel sparkline with visible data only using Aspose.Cells
// Tags: Aspose.Cells sparkline DisplayHidden false | C# sparkline group hide rows columns | Aspose.Cells line sparkline visible data only | Excel sparkline ignore hidden cells Aspose | Aspose.Cells sparkline high low point styling

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills a 5x4 range, hides a specific row and column, adds a line sparkline group covering A1:D5 in column F, sets the group's DisplayHidden property to false so hidden rows and columns are excluded from sparkline calculations, applies high/low point styling, and saves the file as SparklineIgnoreHidden.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (5 rows x 4 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Hide a row (row index 1 => second row) and a column (column index 2 => third column)
            sheet.Cells.Rows[1].IsHidden = true;      // hides row 2
            sheet.Cells.Columns[2].IsHidden = true;   // hides column C

            // Define the area where the sparklines will be placed (one sparkline per row)
            // We'll place them in column F (index 5) rows 0‑4
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a sparkline group of type Line.
            // Data range covers the whole data block A1:D5.
            // The fourth parameter (false) indicates left‑to‑right plotting.
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:D5",
                false,
                sparklineLocation);

            // Retrieve the created group
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Set DisplayHidden to false so hidden rows/columns are ignored
            group.DisplayHidden = false;

            // Optional visual customizations
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;
            group.HighPointColor.Color = Color.Green;
            group.LowPointColor.Color = Color.Red;
            group.LineWeight = 1.0;

            // Save the workbook
            string outputPath = "SparklineIgnoreHidden.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            Console.WriteLine("Hidden rows/columns are ignored in the sparkline calculations.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
