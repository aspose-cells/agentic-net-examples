using System;
using Aspose.Cells;

namespace ScrollBarVisibilityDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // -------------------------------------------------
            // Scenario 1: Show both vertical and horizontal scroll bars
            // -------------------------------------------------
            settings.IsVScrollBarVisible = true;   // Show vertical scroll bar
            settings.IsHScrollBarVisible = true;   // Show horizontal scroll bar

            // Save the workbook for Scenario 1
            workbook.Save("ShowScrollBars.xlsx", SaveFormat.Xlsx);

            // -------------------------------------------------
            // Scenario 2: Hide both vertical and horizontal scroll bars
            // -------------------------------------------------
            settings.IsVScrollBarVisible = false;  // Hide vertical scroll bar
            settings.IsHScrollBarVisible = false;  // Hide horizontal scroll bar

            // Save the workbook for Scenario 2
            workbook.Save("HideScrollBars.xlsx", SaveFormat.Xlsx);

            // -------------------------------------------------
            // Scenario 3: Auto-adjust (use default settings)
            // The default value for both properties is true, which lets Excel
            // display scroll bars as needed based on content.
            // -------------------------------------------------
            settings.IsVScrollBarVisible = true;   // Reset to default
            settings.IsHScrollBarVisible = true;   // Reset to default

            // Save the workbook for Scenario 3
            workbook.Save("AutoAdjustScrollBars.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook files created with different scroll bar visibility settings.");
        }
    }
}