using System;
using Aspose.Cells;

class WorksheetZoomDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the zoom factor (percentage between 10 and 400)
        worksheet.Zoom = 150;

        // Get and display the current zoom factor
        Console.WriteLine("Current worksheet zoom: " + worksheet.Zoom + "%");

        // Save the workbook
        workbook.Save("WorksheetZoomDemo.xlsx");
    }
}