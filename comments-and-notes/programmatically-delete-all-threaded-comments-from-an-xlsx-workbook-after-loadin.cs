using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and clear all comments (including threaded comments)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The Comments collection contains both regular and threaded comments.
            // Calling Clear() removes every comment from the worksheet.
            sheet.Comments.Clear();
        }

        // Save the workbook after removing the threaded comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}