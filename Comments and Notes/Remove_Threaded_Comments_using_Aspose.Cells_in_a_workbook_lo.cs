using Aspose.Cells;
using System;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Clear all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Save the workbook after removing threaded comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}