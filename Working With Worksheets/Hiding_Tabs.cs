using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideWorkbookTabsDemo
    {
        public static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Hide the worksheet tabs at the bottom of the workbook
            workbook.Settings.ShowTabs = false;

            // Save the workbook to an XLSX file
            workbook.Save("HideTabsDemo.xlsx", SaveFormat.Xlsx);

            // Optional: demonstrate re-enabling tabs and saving another file
            workbook.Settings.ShowTabs = true;
            workbook.Save("ShowTabsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}