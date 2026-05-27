using System;
using Aspose.Cells;

class ShowVerticalHideHorizontalScrollBar
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access workbook settings
        WorkbookSettings settings = workbook.Settings;

        // Ensure the vertical scroll bar is visible
        settings.IsVScrollBarVisible = true;

        // Hide the horizontal scroll bar
        settings.IsHScrollBarVisible = false;

        // Save the workbook to disk (XLSX format)
        workbook.Save("VerticalScrollBarOnly.xlsx", SaveFormat.Xlsx);
    }
}