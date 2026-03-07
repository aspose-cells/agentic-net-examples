using System;
using Aspose.Cells;
using System.Drawing;

namespace WorksheetViewsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the view type to Page Break Preview
            worksheet.ViewType = ViewType.PageBreakPreview;

            // Enable ruler visibility (only applicable in Page Break Preview)
            worksheet.IsRulerVisible = true;

            // Set zoom level to 120%
            worksheet.Zoom = 120;

            // Hide gridlines and row/column headers for a cleaner view
            worksheet.IsGridlinesVisible = false;
            worksheet.IsRowColumnHeadersVisible = false;

            // Set the worksheet to display from right to left (optional)
            worksheet.DisplayRightToLeft = false;

            // Add some sample data to illustrate the view
            worksheet.Cells["A1"].PutValue("Worksheet View Demo");
            worksheet.Cells["A2"].PutValue("ViewType: PageBreakPreview");
            worksheet.Cells["A3"].PutValue("Ruler is visible");
            worksheet.Cells["A4"].PutValue("Zoom = 120%");

            // Save the workbook (save rule)
            workbook.Save("WorksheetViewsDemo.xlsx");

            Console.WriteLine("Workbook saved with custom worksheet view settings.");
        }
    }
}