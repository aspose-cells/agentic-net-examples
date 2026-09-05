// Title: Remove password protection from an encrypted Excel workbook and save it as a new file using Aspose.Cells for .NET (C#)
// AI Prompts: Load a password‑protected .xlsx file with Aspose.Cells, call Workbook.Unprotect, and write the workbook to a new location without any password. | Generate C# code that opens an encrypted Excel workbook using LoadOptions.Password, removes protection, and saves an unprotected copy.
// Common Searches: aspnet remove password from encrypted Excel file using Aspose.Cells | c# Aspose.Cells load workbook with password and save without protection | how to unprotect a password‑protected .xlsx programmatically in .NET | save unprotected copy of a protected Excel workbook using Aspose.Cells C#
// Tags: Aspose.Cells unprotect workbook C# | load encrypted Excel with password Aspose.Cells | save unprotected .xlsx Aspose.Cells | Workbook.Unprotect method example | C# remove Excel file password Aspose

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example checks for the protected workbook, loads it with LoadOptions.Password, calls Workbook.Unprotect to clear protection, and then saves the workbook to a new path as an unprotected .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the original protected workbook
                string protectedPath = @"C:\Path\To\ProtectedWorkbook.xlsx";

                if (!File.Exists(protectedPath))
                {
                    Console.WriteLine($"Protected workbook not found: {protectedPath}");
                    return;
                }

                // Load options with the password for encrypted workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = "YourCurrentPassword"
                };

                // Load the workbook using the password
                Workbook workbook = new Workbook(protectedPath, loadOptions);

                // If the workbook has a protection password (not encryption), unprotect it
                workbook.Unprotect("YourCurrentPassword");

                // Path for the new unprotected workbook
                string unprotectedPath = @"C:\Path\To\UnprotectedWorkbook.xlsx";

                // Save the workbook without any password protection
                workbook.Save(unprotectedPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved without protection to: {unprotectedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
