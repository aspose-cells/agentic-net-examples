using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TabBarWidthDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the width of the worksheet tab bar (1/1000 of window width)
            // Example: 1000 means the tab bar occupies the full window width
            workbook.Settings.SheetTabBarWidth = 800; // Adjust as needed

            // Display the current setting
            Console.WriteLine("SheetTabBarWidth: " + workbook.Settings.SheetTabBarWidth);

            // Save the workbook
            workbook.Save("TabBarWidthDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TabBarWidthDemo.Run();
        }
    }
}