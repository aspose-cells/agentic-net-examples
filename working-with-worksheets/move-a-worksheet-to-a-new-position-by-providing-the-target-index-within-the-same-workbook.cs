using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetMoveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets with distinct names
            workbook.Worksheets.Add("First");
            workbook.Worksheets.Add("Second");
            workbook.Worksheets.Add("Third");

            // Retrieve the worksheet to be moved (by name)
            Worksheet sheetToMove = workbook.Worksheets["Third"];

            // Move the worksheet to the desired index (e.g., position 1, which is the second tab)
            sheetToMove.MoveTo(1);

            // Save the workbook to disk
            workbook.Save("MovedWorksheetDemo.xlsx");
        }
    }
}