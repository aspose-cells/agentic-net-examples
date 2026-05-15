using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure narrow page margins (values are in centimeters)
            // These values can be adjusted as needed; here we use 0.5 cm for all sides
            sheet.PageSetup.LeftMargin = 0.5;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;

            // Set the page orientation to Landscape
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Optional: add some data to visualize the effect
            sheet.Cells["A1"].PutValue("Landscape orientation with narrow margins");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
            }

            // Save the workbook (using the provided save rule)
            workbook.Save("NarrowMarginsLandscape.xlsx", SaveFormat.Xlsx);
        }
    }
}