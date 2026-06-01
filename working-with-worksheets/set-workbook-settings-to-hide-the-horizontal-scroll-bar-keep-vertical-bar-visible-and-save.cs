using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access workbook settings
        WorkbookSettings settings = workbook.Settings;

        // Hide the horizontal scroll bar
        settings.IsHScrollBarVisible = false;

        // Ensure the vertical scroll bar remains visible
        settings.IsVScrollBarVisible = true;

        // Save the workbook to an XLSX file
        workbook.Save("WorkbookWithHiddenHScrollBar.xlsx", SaveFormat.Xlsx);
    }
}