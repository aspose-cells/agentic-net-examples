using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ToggleScrollBarVisibility
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Toggle horizontal scroll bar visibility
            settings.IsHScrollBarVisible = !settings.IsHScrollBarVisible;

            // Toggle vertical scroll bar visibility
            settings.IsVScrollBarVisible = !settings.IsVScrollBarVisible;

            // Save the workbook
            workbook.Save("ToggleScrollBars.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Horizontal Scroll Bar Visible: " + settings.IsHScrollBarVisible);
            Console.WriteLine("Vertical Scroll Bar Visible: " + settings.IsVScrollBarVisible);
            Console.WriteLine("Workbook saved as ToggleScrollBars.xlsx");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}