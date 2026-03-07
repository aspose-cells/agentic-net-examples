using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsShowHideDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add additional worksheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Hide the second worksheet using IsVisible property
            workbook.Worksheets[1].IsVisible = false;

            // Set the third worksheet to VeryHidden using VisibilityType
            workbook.Worksheets[2].VisibilityType = VisibilityType.VeryHidden;

            // Hide the workbook tabs
            workbook.Settings.ShowTabs = false;

            // Set the first visible tab index (0‑based)
            workbook.Settings.FirstVisibleTab = 0;

            // Save the workbook with hidden sheets and tabs
            workbook.Save("HiddenDemo.xlsx", SaveFormat.Xlsx);

            // Make everything visible again
            workbook.Settings.ShowTabs = true;
            workbook.Worksheets[1].IsVisible = true;
            workbook.Worksheets[2].VisibilityType = VisibilityType.Visible;

            // Save the workbook with all sheets and tabs visible
            workbook.Save("VisibleDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}