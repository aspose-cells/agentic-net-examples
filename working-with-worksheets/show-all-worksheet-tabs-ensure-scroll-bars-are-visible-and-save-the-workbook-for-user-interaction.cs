using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class ShowTabsAndScrollbars
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add a couple of additional worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Ensure every worksheet is visible (hidden sheets would not have tabs)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsVisible = true;
            }

            // Access the workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Show all worksheet tabs
            settings.ShowTabs = true;

            // Make both vertical and horizontal scrollbars visible
            settings.IsVScrollBarVisible = true;
            settings.IsHScrollBarVisible = true;

            // (Optional) Add some sample data to each sheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells["A1"].PutValue($"Data in {sheet.Name}");
            }

            // Save the workbook so the user can open it and see the tabs and scrollbars
            workbook.Save("ShowAllTabsAndScrollbars.xlsx", SaveFormat.Xlsx);
        }
    }
}