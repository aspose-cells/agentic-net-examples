using System;
using Aspose.Cells;

namespace AsposeCellsScrollBarDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Ensure horizontal scroll bar is visible
            settings.IsHScrollBarVisible = true;

            // Hide the vertical scroll bar
            settings.IsVScrollBarVisible = false;

            // Optionally add some data to visualize the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Horizontal scroll bar visible");
            sheet.Cells["A2"].PutValue("Vertical scroll bar hidden");

            // Save the workbook to disk in XLSX format
            workbook.Save("ScrollBarSettingsDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with horizontal scroll bar shown and vertical scroll bar hidden.");
        }
    }
}