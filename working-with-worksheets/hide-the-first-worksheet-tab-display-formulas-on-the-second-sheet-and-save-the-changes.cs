using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default worksheet is added)
        Workbook workbook = new Workbook();

        // Add a second worksheet so we have at least two sheets
        workbook.Worksheets.Add("Sheet2");

        // Hide the first worksheet (index 0)
        workbook.Worksheets[0].IsVisible = false; // or workbook.Worksheets[0].SetVisible(false, true);

        // Display formulas on the second worksheet (index 1)
        workbook.Worksheets[1].ShowFormulas = true;

        // Save the workbook with the changes
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}