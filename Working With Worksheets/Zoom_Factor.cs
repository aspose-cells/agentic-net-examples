using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the worksheet view zoom factor (percentage)
            sheet.Zoom = 150; // 150% zoom for the UI view

            // Set the page setup zoom factor for printing/exporting
            sheet.PageSetup.Zoom = 150;          // scaling factor in percent
            sheet.PageSetup.IsPercentScale = true; // ensure percent scaling is used

            // Display the current zoom settings
            Console.WriteLine("Worksheet view zoom: " + sheet.Zoom + "%");
            Console.WriteLine("Page setup zoom: " + sheet.PageSetup.Zoom + "%");

            // Save the workbook to a file
            workbook.Save("ZoomDemo.xlsx");
        }
    }
}