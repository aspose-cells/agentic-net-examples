using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsWorksheetMoveAndColor
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (contains a default sheet)
            Workbook workbook = new Workbook();

            // Rename the default sheet and add the remaining sheets
            WorksheetCollection sheets = workbook.Worksheets;
            sheets[0].Name = "Sheet1";
            sheets.Add("Sheet2");
            sheets.Add("Sheet3");

            // Access the worksheet to move (e.g., "Sheet3")
            Worksheet sheetToMove = workbook.Worksheets["Sheet3"];

            // Move the worksheet to index 1 (second position)
            sheetToMove.MoveTo(1);

            // Set the tab color of the moved worksheet to Green
            sheetToMove.TabColor = Color.Green;

            // Save the workbook to a file
            workbook.Save("MovedSheetWithTabColor.xlsx");
        }
    }
}