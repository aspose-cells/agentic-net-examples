using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShowHideScrollBarsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide the workbook's horizontal and vertical scroll bars
        workbook.Settings.IsHScrollBarVisible = false;
        workbook.Settings.IsVScrollBarVisible = false;

        // Add a horizontal scroll bar shape to the worksheet (the shape itself is still visible)
        ScrollBar scrollBar = worksheet.Shapes.AddScrollBar(2, 0, 2, 0, 200, 30);
        scrollBar.IsHorizontal = true;
        scrollBar.Min = 0;
        scrollBar.Max = 100;
        scrollBar.CurrentValue = 25;
        scrollBar.IncrementalChange = 5;
        scrollBar.PageChange = 20;

        // Save the workbook with hidden scroll bars
        workbook.Save("HiddenScrollBars.xlsx", SaveFormat.Xlsx);

        // Load the saved workbook and make the scroll bars visible again
        Workbook loadedWorkbook = new Workbook("HiddenScrollBars.xlsx");
        loadedWorkbook.Settings.IsHScrollBarVisible = true;
        loadedWorkbook.Settings.IsVScrollBarVisible = true;

        // Save the workbook with scroll bars shown
        loadedWorkbook.Save("ShownScrollBars.xlsx", SaveFormat.Xlsx);
    }
}