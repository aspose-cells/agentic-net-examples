using System;
using Aspose.Cells;

namespace AsposeCellsDeleteOptionsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data and a formula that references column A
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1"; // Formula will be affected by column deletion

            // Create DeleteOptions and enable reference updating
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Update formulas and references in other worksheets
            };

            // Delete column A (index 0) using the DeleteOptions instance
            cells.DeleteColumns(0, 1, deleteOptions);

            // After deletion, the formula in C1 should be updated to reference the new column A (originally B)
            Console.WriteLine("Updated formula in C1: " + cells["C1"].Formula);

            // Save the workbook to verify the changes
            workbook.Save("DeleteOptionsDemo.xlsx");
        }
    }
}