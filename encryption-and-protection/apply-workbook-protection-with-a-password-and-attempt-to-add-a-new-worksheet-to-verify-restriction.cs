using System;
using Aspose.Cells;

namespace AsposeCellsWorkbookProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            // This prevents adding, removing, or renaming worksheets
            string password = "mySecretPwd";
            workbook.Protect(ProtectionType.Structure, password);

            // Save the protected workbook
            string protectedPath = "ProtectedWorkbook.xlsx";
            workbook.Save(protectedPath);
            Console.WriteLine($"Workbook saved and protected at: {protectedPath}");

            // Load the protected workbook
            Workbook loadedWorkbook = new Workbook(protectedPath);
            Console.WriteLine("Loaded the protected workbook.");

            // Attempt to add a new worksheet while the workbook is protected
            try
            {
                loadedWorkbook.Worksheets.Add("NewSheetWhileProtected");
                Console.WriteLine("Unexpected: Worksheet added while workbook is protected.");
            }
            catch (Exception ex)
            {
                // Expected exception because the workbook structure is protected
                Console.WriteLine($"Failed to add worksheet as expected: {ex.Message}");
            }

            // Unprotect the workbook using the correct password
            loadedWorkbook.Unprotect(password);
            Console.WriteLine("Workbook unprotected successfully.");

            // Now adding a new worksheet should succeed
            try
            {
                loadedWorkbook.Worksheets.Add("NewSheetAfterUnprotect");
                Console.WriteLine("Worksheet added after unprotecting the workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error after unprotecting: {ex.Message}");
            }

            // Save the final workbook
            string finalPath = "UnprotectedWorkbook.xlsx";
            loadedWorkbook.Save(finalPath);
            Console.WriteLine($"Final workbook saved at: {finalPath}");
        }
    }
}