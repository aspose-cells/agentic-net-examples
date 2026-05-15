using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineFirstLastPointDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the sparkline (row 4, columns A‑D)
            sheet.Cells["A4"].PutValue(5);
            sheet.Cells["B4"].PutValue(2);
            sheet.Cells["C4"].PutValue(8);
            sheet.Cells["D4"].PutValue(3);

            // Define the location of the sparkline: cell I4
            // Column I = index 8 (0‑based), Row 4 = index 3
            CellArea location = new CellArea
            {
                StartColumn = 8,
                EndColumn = 8,
                StartRow = 3,
                EndRow = 3
            };

            // Add a line sparkline group with the data range A4:D4 and place it at I4
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A4:D4", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Enable highlighting of the first and last points
            group.ShowFirstPoint = true;
            group.ShowLastPoint = true;

            // Optional: set colors for the first and last points
            CellsColor firstColor = workbook.CreateCellsColor();
            firstColor.Color = Color.Purple;
            group.FirstPointColor = firstColor;

            CellsColor lastColor = workbook.CreateCellsColor();
            lastColor.Color = Color.Yellow;
            group.LastPointColor = lastColor;

            // Save the workbook
            workbook.Save("SparklineFirstLastPointDemo.xlsx");
        }
    }
}