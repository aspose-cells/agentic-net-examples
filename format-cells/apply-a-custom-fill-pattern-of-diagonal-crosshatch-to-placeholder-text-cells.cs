using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsPatternDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell that contains placeholder text
            Cell placeholderCell = worksheet.Cells["A1"];
            placeholderCell.PutValue("Placeholder Text");

            // Retrieve the cell's current style
            Style style = placeholderCell.GetStyle();

            // Apply diagonal crosshatch background pattern
            style.Pattern = BackgroundType.DiagonalCrosshatch;

            // Set foreground (pattern) and background colors
            style.ForegroundColor = Color.Black;   // Color of the crosshatch lines
            style.BackgroundColor = Color.White;   // Color behind the pattern

            // Apply the modified style back to the cell
            placeholderCell.SetStyle(style);

            // Save the workbook
            workbook.Save("PlaceholderPattern.xlsx", SaveFormat.Xlsx);
        }
    }
}