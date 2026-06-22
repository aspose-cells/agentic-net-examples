using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class WorkbookStructureProtectionDemo
    {
        static void Main()
        {
            // Paths
            string inputPath = "ProtectedWorkbook.xlsx";
            string outputPath = "WorkbookStructureRenamed.xlsx";

            // Passwords
            string oldPassword = "oldPassword123";
            string newPassword = "newPassword456";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook inside using to ensure disposal
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Attempt to unprotect workbook structure with the old password
                    try
                    {
                        workbook.Unprotect(oldPassword);
                    }
                    catch (Exception ex)
                    {
                        // If the password is invalid, Aspose.Cells throws a generic exception.
                        // We simply log and continue without aborting.
                        Console.WriteLine($"Unprotect failed (likely invalid password): {ex.Message}");
                    }

                    // Rename first worksheet
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "RenamedSheet";

                    // Apply structure protection with new password
                    workbook.Protect(ProtectionType.Structure, newPassword);

                    // Save the modified workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                }

                Console.WriteLine($"Workbook saved to '{outputPath}' with new structure password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}