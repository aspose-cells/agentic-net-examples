using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MakeWorksheetVisibleDemo
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add two additional worksheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Hide the second worksheet (index 1)
            workbook.Worksheets[1].IsVisible = false;

            // Make the hidden worksheet visible again
            workbook.Worksheets[1].IsVisible = true;

            // Save the workbook to a file
            workbook.Save("MakeWorksheetVisibleDemo.xlsx");
        }
    }
}