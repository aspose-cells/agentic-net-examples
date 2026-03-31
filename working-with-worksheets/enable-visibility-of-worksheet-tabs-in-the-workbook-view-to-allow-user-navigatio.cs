using System;
using Aspose.Cells;

namespace EnableWorksheetTabsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Ensure worksheet tabs are visible in the workbook view
            workbook.Settings.ShowTabs = true;

            // Optionally, you can set the first visible tab index if needed
            // workbook.Settings.FirstVisibleTab = 0;

            // Save the workbook to a file
            workbook.Save("WorkbookWithTabsVisible.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with worksheet tabs visible.");
        }
    }
}