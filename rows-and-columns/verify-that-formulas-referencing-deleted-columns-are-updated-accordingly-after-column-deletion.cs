using System;
using Aspose.Cells;

namespace AsposeCellsColumnDeletionVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in columns A and B
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1"; // Formula references columns A and B

            cells["A2"].PutValue(30);
            cells["B2"].PutValue(40);
            cells["C2"].Formula = "=A2+B2";

            // Display formulas before deletion
            Console.WriteLine("Formulas before column deletion:");
            Console.WriteLine($"C1: {cells["C1"].Formula}");
            Console.WriteLine($"C2: {cells["C2"].Formula}");

            // Set up DeleteOptions to update references after deletion
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Ensure formulas are adjusted
            };

            // Delete column A (index 0). After deletion, original column B becomes column A.
            cells.DeleteColumns(0, 1, deleteOptions);

            // Display formulas after deletion to verify they have been updated
            Console.WriteLine("\nFormulas after column deletion:");
            Console.WriteLine($"C1: {cells["C1"].Formula}"); // Expected: "=A1+B1" becomes "=A1+B1" where former B1 is now A1
            Console.WriteLine($"C2: {cells["C2"].Formula}");

            // Save the workbook (optional, demonstrates usage of save rule)
            workbook.Save("ColumnDeletionVerification.xlsx");
        }
    }
}