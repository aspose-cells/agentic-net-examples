using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 0 through 9 (10 rows total)
        worksheet.Cells.HideRows(0, 10);

        // Hide the vertical scroll bar (toggle visibility)
        workbook.Settings.IsVScrollBarVisible = false;

        // Save the workbook to a file
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}