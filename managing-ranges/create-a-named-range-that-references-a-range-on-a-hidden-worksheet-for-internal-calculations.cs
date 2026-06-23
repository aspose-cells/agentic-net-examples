using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeHiddenSheet
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will be used for internal calculations
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");

            // Hide the worksheet so it is not visible to end users
            hiddenSheet.IsVisible = false;

            // Populate some sample data on the hidden sheet (range A1:B5)
            for (int row = 0; row < 5; row++)
            {
                hiddenSheet.Cells[row, 0].PutValue(row + 1);          // Column A
                hiddenSheet.Cells[row, 1].PutValue((row + 1) * 10);   // Column B
            }

            // Create a named range that refers to the range on the hidden worksheet
            int nameIndex = workbook.Worksheets.Names.Add("CalcRange");
            Name calcRange = workbook.Worksheets.Names[nameIndex];
            // The RefersTo formula must start with an equal sign
            calcRange.RefersTo = $"=HiddenSheet!$A$1:$B$5";

            // Optionally make the name itself invisible (not shown in Name Manager)
            calcRange.IsVisible = false;

            // Save the workbook to a file
            workbook.Save("WorkbookWithHiddenNamedRange.xlsx");
        }
    }
}