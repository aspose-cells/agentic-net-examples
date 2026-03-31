using System;
using Aspose.Cells;

namespace AsposeCellsVisibilityToggle
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // -------------------------------------------------
            // Hide workbook tabs (the tab bar at the bottom)
            // -------------------------------------------------
            workbook.Settings.ShowTabs = false;

            // Hide the second worksheet (index 1) using IsVisible property
            workbook.Worksheets[1].IsVisible = false;

            // Hide the third worksheet (index 2) using VisibilityType for demonstration
            workbook.Worksheets[2].VisibilityType = VisibilityType.Hidden;

            // Save the workbook with hidden tabs and hidden sheets
            workbook.Save("Workbook_Hidden.xlsx", SaveFormat.Xlsx);

            // -------------------------------------------------
            // Toggle visibility back to visible
            // -------------------------------------------------
            // Show workbook tabs again
            workbook.Settings.ShowTabs = true;

            // Make all worksheets visible
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsVisible = true;               // Simple visibility flag
                sheet.VisibilityType = VisibilityType.Visible; // Ensure enum state is also Visible
            }

            // Save the workbook with all tabs and sheets visible
            workbook.Save("Workbook_Visible.xlsx", SaveFormat.Xlsx);
        }
    }
}