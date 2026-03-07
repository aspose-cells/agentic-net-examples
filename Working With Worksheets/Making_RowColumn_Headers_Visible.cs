using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RowColumnHeadersVisibilityDemo
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
            string hiddenPath = "HeadersHidden.xlsx";
            workbook.Save(hiddenPath);

            // Load the saved workbook to verify the setting
            Workbook loadedWorkbook = new Workbook(hiddenPath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Output the visibility status (expected: False)
            Console.WriteLine("Row and Column Headers Visible after load: " + loadedWorksheet.IsRowColumnHeadersVisible);

            // Make the headers visible again
            loadedWorksheet.IsRowColumnHeadersVisible = true;

            // Save the workbook with headers visible
            string visiblePath = "HeadersVisible.xlsx";
            loadedWorkbook.Save(visiblePath);
        }
    }
}