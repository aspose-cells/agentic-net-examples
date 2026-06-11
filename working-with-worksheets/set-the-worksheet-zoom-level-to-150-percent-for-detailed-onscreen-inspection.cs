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

            // Set the zoom level to 150%
            worksheet.Zoom = 150;

            // Save the workbook to a file
            workbook.Save("WorksheetZoom150.xlsx");

            // Optional: output confirmation
            Console.WriteLine("Worksheet zoom set to 150% and saved as WorksheetZoom150.xlsx");
        }
    }
}