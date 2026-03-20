using System;
using Aspose.Cells;

namespace AsposeCellsOrientationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add(); // Sheet2
            workbook.Worksheets.Add(); // Sheet3

            // -------------------------------------------------
            // Set orientation for the first worksheet (index 0) to Landscape
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.PageSetup.Orientation = PageOrientationType.Landscape;

            // Set orientation for the second worksheet (index 1) to Portrait
            Worksheet sheet2 = workbook.Worksheets[1];
            sheet2.PageSetup.Orientation = PageOrientationType.Portrait;

            // -------------------------------------------------
            // Set orientation for all worksheets in the workbook to Portrait
            // using the WorkbookSettings method
            workbook.Settings.SetPageOrientationType(PageOrientationType.Portrait);

            // -------------------------------------------------
            // Retrieve and display the orientation of each worksheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                PageOrientationType orientation = ws.PageSetup.Orientation;
                Console.WriteLine($"Worksheet \"{ws.Name}\" orientation: {orientation}");
            }

            // -------------------------------------------------
            // Save the workbook to demonstrate that the settings are persisted
            workbook.Save("OrientationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}