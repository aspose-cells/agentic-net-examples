using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and clear all comments (standard and threaded)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments(); // Removes every comment from the worksheet
        }

        // Save the workbook after comments have been removed
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}