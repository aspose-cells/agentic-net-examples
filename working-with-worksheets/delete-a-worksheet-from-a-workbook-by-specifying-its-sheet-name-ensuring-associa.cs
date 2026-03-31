using System;
using Aspose.Cells;

class DeleteWorksheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add worksheets to the workbook
        workbook.Worksheets.Add("KeepSheet");
        workbook.Worksheets.Add("DeleteMe");

        // Populate the sheet that will be removed (optional)
        workbook.Worksheets["DeleteMe"].Cells["A1"].PutValue("This data will be removed");

        // Delete the worksheet by specifying its name
        workbook.Worksheets.RemoveAt("DeleteMe");

        // Save the modified workbook
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}