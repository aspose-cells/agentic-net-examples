using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ScrollBarVisibilityDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Ensure both horizontal and vertical scroll bars are visible
            settings.IsHScrollBarVisible = true; // Show horizontal scroll bar
            settings.IsVScrollBarVisible = true; // Show vertical scroll bar

            // Save the workbook
            workbook.Save("ScrollBarVisibilityDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with both scroll bars visible.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ScrollBarVisibilityDemo.Run();
        }
    }
}