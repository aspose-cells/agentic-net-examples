using System;
using Aspose.Cells;

namespace HideWorksheetTabsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Hide the worksheet tabs to prevent user navigation
            workbook.Settings.ShowTabs = false;

            // Save the workbook to a file (XLSX format)
            workbook.Save("HideTabsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}