using System;
using Aspose.Cells;

namespace WorkbookUnprotectAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the protected and the resulting unprotected workbook
            string protectedFilePath = "ProtectedWorkbook.xlsx";
            string unprotectedFilePath = "UnprotectedWorkbook.xlsx";

            // Recovered password for the workbook structure protection
            string password = "recoveredPassword";

            try
            {
                // Load the protected workbook
                Workbook workbook = new Workbook(protectedFilePath);

                // Check if the workbook is actually protected with a password
                if (workbook.IsWorkbookProtectedWithPassword)
                {
                    // Unprotect the workbook using the recovered password
                    workbook.Unprotect(password);

                    // Log the successful unprotection event
                    Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook '{protectedFilePath}' was unprotected using the provided password.");
                }
                else
                {
                    // Log that the workbook was not password‑protected
                    Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook '{protectedFilePath}' is not protected with a password; no unprotection needed.");
                }

                // Save the unprotected workbook
                workbook.Save(unprotectedFilePath);
                Console.WriteLine($"Workbook saved as '{unprotectedFilePath}'.");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during the process
                Console.WriteLine($"[{DateTime.UtcNow:u}] Error unprotecting workbook: {ex.Message}");
            }
        }
    }
}