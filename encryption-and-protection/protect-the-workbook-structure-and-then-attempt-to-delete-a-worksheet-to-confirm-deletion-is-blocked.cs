using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default has one worksheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            // This prevents adding, removing, renaming, or moving worksheets
            workbook.Protect(ProtectionType.Structure, "myPassword");

            // Verify that the workbook is protected
            Console.WriteLine("Workbook structure protected: " + workbook.Settings.IsProtected);
            Console.WriteLine("Workbook protected with password: " + workbook.IsWorkbookProtectedWithPassword);

            // Attempt to delete the first worksheet; should raise an exception
            try
            {
                workbook.Worksheets.RemoveAt(0);
                Console.WriteLine("Worksheet removed successfully (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to remove worksheet as expected: " + ex.Message);
            }

            // Save the workbook (protected)
            workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}