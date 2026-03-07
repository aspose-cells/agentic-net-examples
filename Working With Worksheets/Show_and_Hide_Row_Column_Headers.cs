using System;
using Aspose.Cells;

namespace ShowHideRowColumnHeadersDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide row and column headers
            worksheet.IsRowColumnHeadersVisible = false;

            // Save the workbook with headers hidden
            workbook.Save("HeadersHidden.xlsx");

            // Load the saved workbook to verify the property
            Workbook loadedWorkbook = new Workbook("HeadersHidden.xlsx");
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Headers visible after hide: " + loadedWorksheet.IsRowColumnHeadersVisible);

            // Show row and column headers again
            loadedWorksheet.IsRowColumnHeadersVisible = true;

            // Save the workbook with headers visible
            loadedWorkbook.Save("HeadersVisible.xlsx");

            // Load the final workbook to confirm the change
            Workbook finalWorkbook = new Workbook("HeadersVisible.xlsx");
            Console.WriteLine("Headers visible after show: " + finalWorkbook.Worksheets[0].IsRowColumnHeadersVisible);
        }
    }
}