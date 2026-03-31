using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet zoom factor to 150%
            worksheet.Zoom = 150; // uses Worksheet.Zoom property

            // Retrieve and display the current zoom factor
            int currentZoom = worksheet.Zoom;
            Console.WriteLine("Current worksheet zoom: " + currentZoom + "%");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("WorksheetZoomDemo.xlsx");
        }
    }
}