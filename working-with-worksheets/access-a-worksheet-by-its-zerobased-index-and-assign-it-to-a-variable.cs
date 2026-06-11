using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetAccess
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (contains at least one default worksheet)
            Workbook workbook = new Workbook();

            // Access the worksheet at zero‑based index 0 and assign it to a variable
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Example operation: set the worksheet name
            firstWorksheet.Name = "FirstSheet";

            // Save the workbook to verify that the code runs without errors
            workbook.Save("WorksheetAccessDemo.xlsx");
        }
    }
}