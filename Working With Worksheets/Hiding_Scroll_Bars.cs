using System;
using Aspose.Cells;

class HideScrollBarsDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access workbook settings and hide both scroll bars
        workbook.Settings.IsHScrollBarVisible = false; // hide horizontal scroll bar
        workbook.Settings.IsVScrollBarVisible = false; // hide vertical scroll bar

        // Save the workbook to a file
        workbook.Save("HiddenScrollBars.xlsx", SaveFormat.Xlsx);
    }
}