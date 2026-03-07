using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ScrollBarVisibilityDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that both horizontal and vertical scroll bars are visible in the generated file
            workbook.Settings.IsHScrollBarVisible = true; // Show horizontal scroll bar
            workbook.Settings.IsVScrollBarVisible = true; // Show vertical scroll bar

            // Add a horizontal scroll bar shape to the worksheet (optional visual demonstration)
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            ScrollBar hScrollBar = worksheet.Shapes.AddScrollBar(5, 0, 2, 0, 20, 200);
            hScrollBar.IsHorizontal = true;
            hScrollBar.Min = 0;
            hScrollBar.Max = 100;
            hScrollBar.CurrentValue = 30;
            hScrollBar.IncrementalChange = 5;
            hScrollBar.PageChange = 20;

            // Add a vertical scroll bar shape to the worksheet (optional visual demonstration)
            ScrollBar vScrollBar = worksheet.Shapes.AddScrollBar(10, 0, 5, 0, 200, 20);
            vScrollBar.IsHorizontal = false;
            vScrollBar.Min = 0;
            vScrollBar.Max = 100;
            vScrollBar.CurrentValue = 70;
            vScrollBar.IncrementalChange = 5;
            vScrollBar.PageChange = 20;

            // Save the workbook to an XLSX file
            workbook.Save("ScrollBarsVisibleDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with scroll bars visible.");
        }
    }
}