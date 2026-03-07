using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideRowColumnHeadersDemo
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide the row and column headers (property from Worksheet)
            worksheet.IsRowColumnHeadersVisible = false;

            // Save the workbook to verify the setting (lifecycle: save)
            workbook.Save("HideHeadersOutput.xlsx");

            // Load the saved workbook to confirm the property (lifecycle: load)
            Workbook loadedWorkbook = new Workbook("HideHeadersOutput.xlsx");
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Output the current visibility status of row/column headers
            Console.WriteLine("Row and Column Headers Visible: " + loadedWorksheet.IsRowColumnHeadersVisible);
        }
    }
}