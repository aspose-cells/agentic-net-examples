using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ShowHideTabsDemo
    {
        public static void Run()
        {
            // Create a new workbook and hide the worksheet tabs
            Workbook workbookHide = new Workbook();
            workbookHide.Settings.ShowTabs = false; // hide tabs
            workbookHide.Save("HideTabsDemo.xlsx", SaveFormat.Xlsx);

            // Create another workbook and ensure the worksheet tabs are visible (default is true)
            Workbook workbookShow = new Workbook();
            workbookShow.Settings.ShowTabs = true; // show tabs
            workbookShow.Save("ShowTabsDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShowHideTabsDemo.Run();
        }
    }
}