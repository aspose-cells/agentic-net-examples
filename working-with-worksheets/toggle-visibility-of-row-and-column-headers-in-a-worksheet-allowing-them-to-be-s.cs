using System;
using Aspose.Cells;

namespace WorksheetHeaderVisibilityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide the row and column headers (feature: IsRowColumnHeadersVisible)
            worksheet.IsRowColumnHeadersVisible = false;

            // Save the workbook to a file (lifecycle: save)
            string filePath = "HeadersHidden.xlsx";
            workbook.Save(filePath);

            // Load the saved workbook to verify the setting (lifecycle: load)
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Output the current visibility state of the headers
            Console.WriteLine("Row and Column Headers Visible: " + loadedWorksheet.IsRowColumnHeadersVisible);
        }
    }
}