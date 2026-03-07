using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            WorkbookTabsVisibilityDemo.Run();
        }
    }

    public class WorkbookTabsVisibilityDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            workbook.Settings.ShowTabs = true;
            workbook.Save("WorkbookWithVisibleTabs.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved with tabs visible.");
        }
    }
}