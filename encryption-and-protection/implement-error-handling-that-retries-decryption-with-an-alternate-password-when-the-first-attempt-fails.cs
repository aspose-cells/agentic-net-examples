// Title: Open an encrypted .xlsx workbook with a primary password and automatically retry using an alternate password in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a password‑protected Excel file with Aspose.Cells LoadOptions, catches CellsException, and retries the load using a second password. | Add logic to write the successfully opened workbook to a new .xlsx file without any password after either password succeeds. | Include a file‑existence check and clear console messages for both password attempts and for any save errors.
// Common Searches: Aspose.Cells load encrypted Excel file with fallback password C# | Retry opening a password‑protected workbook using a second password in .NET | Decrypt .xlsx when first password fails using Aspose.Cells | Save unprotected copy of an encrypted workbook after successful decryption Aspose.Cells
// Tags: fallback password handling Aspose.Cells | loadoptions password retry .NET | decrypt encrypted workbook alternate password C# | save workbook without password Aspose.Cells | exception handling encrypted Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The program checks for an encrypted Excel file, attempts to open it with a primary password, retries with an alternate password if the first fails, and saves an unprotected copy of the workbook.
    class DecryptWorkbook
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // First and alternate passwords
            string primaryPassword = "FirstPassword";
            string alternatePassword = "AlternatePassword";

            // Verify that the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Attempt to load with the primary password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = primaryPassword };
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (CellsException)
            {
                try
                {
                    // Primary password failed – retry with the alternate password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = alternatePassword };
                    workbook = new Workbook(filePath, loadOptions);
                }
                catch (CellsException ex)
                {
                    Console.WriteLine($"Failed to open workbook with both passwords: {ex.Message}");
                    return;
                }
            }

            try
            {
                // Save a copy without password
                workbook.Save("decrypted_copy.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as decrypted_copy.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
