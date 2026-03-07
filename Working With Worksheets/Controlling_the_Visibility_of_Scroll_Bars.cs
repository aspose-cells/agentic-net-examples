using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsScrollBarVisibilityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access workbook settings to control the visibility of the worksheet scroll bars
            // Hide the horizontal scroll bar
            workbook.Settings.IsHScrollBarVisible = false;
            // Show the vertical scroll bar (default is true, set explicitly for clarity)
            workbook.Settings.IsVScrollBarVisible = true;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a horizontal scroll bar shape to the worksheet
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            ScrollBar hScrollBar = sheet.Shapes.AddScrollBar(2, 0, 2, 0, 20, 200);
            hScrollBar.IsHorizontal = true;          // Ensure it is horizontal
            hScrollBar.Min = 0;
            hScrollBar.Max = 100;
            hScrollBar.CurrentValue = 30;
            hScrollBar.IncrementalChange = 5;
            hScrollBar.PageChange = 20;

            // Add a vertical scroll bar shape to the worksheet
            ScrollBar vScrollBar = sheet.Shapes.AddScrollBar(5, 0, 5, 0, 200, 20);
            vScrollBar.IsHorizontal = false;         // Make it vertical
            vScrollBar.Min = 0;
            vScrollBar.Max = 100;
            vScrollBar.CurrentValue = 70;
            vScrollBar.IncrementalChange = 10;
            vScrollBar.PageChange = 30;

            // Save the workbook (save rule)
            workbook.Save("ScrollBarVisibilityDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook created with custom scroll bar visibility settings.");
        }
    }
}