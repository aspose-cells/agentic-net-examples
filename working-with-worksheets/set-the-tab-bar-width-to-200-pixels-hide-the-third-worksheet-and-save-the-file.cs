using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the width of the worksheet tab bar (value is in 1/1000 of window width)
        workbook.Settings.SheetTabBarWidth = 200; // Approximate width

        // Add two more worksheets so we have at least three sheets
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Hide the third worksheet (index 2, zero‑based)
        workbook.Worksheets[2].IsVisible = false; // alternatively: SetVisible(false, true)

        // Save the workbook to a file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}