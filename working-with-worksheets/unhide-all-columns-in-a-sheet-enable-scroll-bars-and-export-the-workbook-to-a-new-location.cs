using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Unhide a large range of columns (e.g., first 1000 columns) and set a default width
        cells.UnhideColumns(0, 1000, 10.0);

        // Ensure both horizontal and vertical scroll bars are visible
        workbook.Settings.IsHScrollBarVisible = true;
        workbook.Settings.IsVScrollBarVisible = true;

        // Save the workbook to a new location
        workbook.Save(@"C:\Temp\UnhiddenWorkbook.xlsx", SaveFormat.Xlsx);
    }
}