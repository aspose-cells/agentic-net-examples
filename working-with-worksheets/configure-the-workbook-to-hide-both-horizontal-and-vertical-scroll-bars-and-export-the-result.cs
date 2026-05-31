using System;
using Aspose.Cells;

namespace AsposeCellsScrollBarDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Hide horizontal and vertical scroll bars
            settings.IsHScrollBarVisible = false;
            settings.IsVScrollBarVisible = false;

            // Save the workbook to an XLSX file
            workbook.Save("HiddenScrollBars.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with both scroll bars hidden.");
        }
    }
}