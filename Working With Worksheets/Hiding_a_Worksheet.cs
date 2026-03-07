using System;
using Aspose.Cells;

class HideWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add additional worksheets
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Hide the added worksheets using SetVisible method
        // First parameter: false to hide, second parameter: true to ignore errors if hiding is not allowed
        workbook.Worksheets["Sheet2"].SetVisible(false, true);
        workbook.Worksheets["Sheet3"].SetVisible(false, true);

        // Optionally, hide the first worksheet using the IsVisible property
        // workbook.Worksheets[0].IsVisible = false;

        // Save the workbook with hidden sheets
        workbook.Save("HiddenSheetsDemo.xlsx", SaveFormat.Xlsx);
    }
}