using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Add sample worksheets
        workbook.Worksheets.Add("SheetA");
        workbook.Worksheets.Add("SheetB");
        workbook.Worksheets.Add("SheetC");

        // Select the worksheet to move (e.g., "SheetC")
        Worksheet sheetToMove = workbook.Worksheets["SheetC"];

        // Move the worksheet to the desired index (e.g., position 1)
        sheetToMove.MoveTo(1); // Moves "SheetC" to be the second tab

        // Update the tab color of the moved worksheet
        sheetToMove.TabColor = Color.LightBlue;

        // Save the workbook (lifecycle save)
        workbook.Save("MovedAndColoredSheet.xlsx");
    }
}