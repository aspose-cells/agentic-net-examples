using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the zoom factor to 150%
            worksheet.Zoom = 150;

            // Output the current zoom factor for verification
            Console.WriteLine($"Worksheet zoom set to {worksheet.Zoom}%");

            // Save the workbook
            workbook.Save("WorksheetZoom150.xlsx");
        }
    }
}