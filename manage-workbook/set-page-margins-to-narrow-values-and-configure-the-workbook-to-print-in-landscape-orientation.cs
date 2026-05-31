using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set narrow page margins (values in centimeters)
            // Adjust these values as needed for "narrow" margins
            sheet.PageSetup.TopMargin = 0.5;    // 0.5 cm top margin
            sheet.PageSetup.BottomMargin = 0.5; // 0.5 cm bottom margin
            sheet.PageSetup.LeftMargin = 0.5;   // 0.5 cm left margin
            sheet.PageSetup.RightMargin = 0.5;  // 0.5 cm right margin

            // Configure the worksheet to print in landscape orientation
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Add some sample data to visualize the effect
            sheet.Cells["A1"].PutValue("Landscape orientation with narrow margins");
            for (int row = 2; row <= 20; row++)
            {
                sheet.Cells[$"A{row}"].PutValue($"Row {row - 1}");
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("NarrowMarginsLandscape.xlsx", SaveFormat.Xlsx);
        }
    }
}