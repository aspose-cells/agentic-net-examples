using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the width of the worksheet tab bar (in 1/1000 of window width)
        workbook.Settings.SheetTabBarWidth = 1000; // example value

        // Ensure there are at least four worksheets
        while (workbook.Worksheets.Count < 4)
        {
            workbook.Worksheets.Add();
        }

        // Hide the fourth worksheet (zero‑based index 3)
        workbook.Worksheets[3].SetVisible(false, true);

        // Save the workbook with the applied settings
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}