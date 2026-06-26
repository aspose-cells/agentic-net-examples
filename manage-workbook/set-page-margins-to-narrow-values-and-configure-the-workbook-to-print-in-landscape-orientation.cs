using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the PageSetup object of the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Set narrow margins (values are in centimeters)
            pageSetup.LeftMargin = 0.5;   // 0.5 cm left margin
            pageSetup.RightMargin = 0.5;  // 0.5 cm right margin
            pageSetup.TopMargin = 0.5;    // 0.5 cm top margin
            pageSetup.BottomMargin = 0.5; // 0.5 cm bottom margin

            // Configure the worksheet to print in landscape orientation
            pageSetup.Orientation = PageOrientationType.Landscape;

            // (Optional) Add some sample data to visualize the effect
            sheet.Cells["A1"].PutValue("Landscape orientation with narrow margins");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
            }

            // Save the workbook to a file
            workbook.Save("NarrowMargins_Landscape.xlsx", SaveFormat.Xlsx);
        }
    }
}