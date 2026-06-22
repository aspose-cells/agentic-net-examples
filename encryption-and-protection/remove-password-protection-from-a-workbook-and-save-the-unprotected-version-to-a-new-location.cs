using System;
using System.IO;
using Aspose.Cells;

class RemoveWorkbookPassword
{
    static void Main()
    {
        try
        {
            // Path to the password‑protected workbook
            string protectedPath = "protected_workbook.xlsx";

            // Password used to protect the workbook
            string password = "myPassword";

            // Verify that the source file exists
            if (!File.Exists(protectedPath))
            {
                Console.WriteLine($"File not found: {protectedPath}");
                return;
            }

            Workbook workbook = null;

            // Attempt to load the workbook with the supplied password (for encrypted files)
            try
            {
                var loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(protectedPath, loadOptions);
            }
            catch (Exception)
            {
                // If loading with a password fails (e.g., not encrypted), load without a password
                workbook = new Workbook(protectedPath);
            }

            // Remove workbook protection (if any)
            try
            {
                workbook.Unprotect(password);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unprotect failed: {ex.Message}");
            }

            // Remove encryption password before saving
            workbook.Settings.Password = string.Empty;

            // Save the unprotected workbook to a new file
            string unprotectedPath = "unprotected_workbook.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(unprotectedPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(unprotectedPath);
            Console.WriteLine($"Workbook saved without password to: {unprotectedPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}