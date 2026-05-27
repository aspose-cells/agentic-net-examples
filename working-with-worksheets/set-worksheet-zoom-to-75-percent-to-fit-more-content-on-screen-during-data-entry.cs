using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the zoom factor to 75%
            worksheet.Zoom = 75;

            // Optionally display the current zoom factor
            Console.WriteLine("Worksheet zoom set to: " + worksheet.Zoom + "%");

            // Save the workbook to a file
            workbook.Save("WorksheetZoom75.xlsx");
        }
    }
}