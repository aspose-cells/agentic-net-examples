using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SheetTabBarWidthDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the width of the worksheet tab bar (value is in 1/1000 of window width)
            workbook.Settings.SheetTabBarWidth = 1200; // Example: 1200 = 1.2 times the window width

            // Retrieve and display the current SheetTabBarWidth value
            Console.WriteLine("SheetTabBarWidth: " + workbook.Settings.SheetTabBarWidth);

            // Save the workbook to a file
            workbook.Save("SheetTabBarWidthDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SheetTabBarWidthDemo.Run();
        }
    }
}