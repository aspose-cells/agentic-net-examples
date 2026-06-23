using System;
using Aspose.Cells;

namespace AsposeCellsActiveSheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add two more worksheets so we have at least three sheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Set the active worksheet to the third sheet (zero‑based index 2)
            workbook.Worksheets.ActiveSheetIndex = 2;

            // Optional: verify the active sheet name
            Console.WriteLine("Active Sheet: " + workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex].Name);

            // Save the workbook to a file
            workbook.Save("ActiveSheetSet.xlsx");
        }
    }
}