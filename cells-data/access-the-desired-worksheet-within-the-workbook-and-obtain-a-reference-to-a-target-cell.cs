using System;
using Aspose.Cells;

namespace AsposeCellsAccessCellDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain a reference to a target cell (e.g., cell B2)
            Cell targetCell = worksheet.Cells["B2"];

            // Example: put a value into the target cell
            targetCell.PutValue("Hello Aspose!");

            // Save the workbook to verify the changes (lifecycle: save)
            workbook.Save("AccessCellDemo.xlsx");
        }
    }
}