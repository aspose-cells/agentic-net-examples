using System;
using Aspose.Cells;

namespace AsposeCellsTabVisibilityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Hide the worksheet tabs
            workbook.Settings.ShowTabs = false;

            // Save the workbook with tabs hidden
            workbook.Save("HideTabsDemo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved with worksheet tabs hidden.");

            // Show the worksheet tabs (default is true)
            workbook.Settings.ShowTabs = true;

            // Save the workbook with tabs visible
            workbook.Save("ShowTabsDemo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved with worksheet tabs visible.");
        }
    }
}