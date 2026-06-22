using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookStructureProtectionDemo
    {
        // Entry point required for console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Protect the workbook structure with an initial password
            string oldPassword = "oldPass";
            workbook.Protect(ProtectionType.Structure, oldPassword);

            // Verify that the workbook is protected with a password
            Console.WriteLine("Initially protected with password: " + workbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("Workbook settings indicate protection: " + workbook.Settings.IsProtected);

            // Unprotect the workbook using the old password
            workbook.Unprotect(oldPassword);

            // Verify that the workbook is no longer protected
            Console.WriteLine("After unprotect, protected with password: " + workbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("After unprotect, settings indicate protection: " + workbook.Settings.IsProtected);

            // Re‑apply protection with a new password
            string newPassword = "newPass";
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Verify that the new protection is in effect
            Console.WriteLine("Re‑protected with new password: " + workbook.IsWorkbookProtectedWithPassword);
            Console.WriteLine("Workbook settings indicate protection after re‑protect: " + workbook.Settings.IsProtected);

            // Save the workbook (lifecycle: save)
            string outputPath = "ReprotectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}