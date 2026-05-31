using System;
using Aspose.Cells;

class ShowFormulasHiddenSheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet and hide it
        Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
        hiddenSheet.IsVisible = false; // hide the worksheet

        // Add a formula to the hidden worksheet
        hiddenSheet.Cells["A1"].Formula = "=1+2+3";

        // Make the hidden worksheet visible again
        hiddenSheet.IsVisible = true; // or hiddenSheet.SetVisible(true, true);

        // Set the worksheet to display formulas instead of calculated values
        hiddenSheet.ShowFormulas = true;

        // Save the workbook to a file
        workbook.Save("ShowFormulasDemo.xlsx");
    }
}