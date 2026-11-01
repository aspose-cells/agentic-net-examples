// Title: Decrypt a password‑protected XLSX workbook and recalculate its formulas using Aspose.Cells for .NET
// Common Searches: C# how to open an encrypted XLSX file with a password using Aspose.Cells | Aspose.Cells load protected Excel workbook and force formula recalculation | Decrypt Excel file and save as new workbook after calculating formulas in .NET | Verify formula result after opening password‑protected workbook with Aspose.Cells | LoadOptions password example for encrypted XLSX in C#
// Tags: encrypted xlsx loading password Aspose.Cells | workbook formula recalculation Aspose.Cells | decrypted workbook saving xlsx Aspose.Cells | cell A1 formula value retrieval Aspose.Cells | protected excel file error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example demonstrates how to open a password‑protected XLSX file using LoadOptions with the correct password, recalculate all formulas via Workbook.CalculateFormula, read a specific formula cell value for verification, and optionally save the decrypted workbook as a new XLSX file, with proper exception handling.
    class DecryptAndRecalculate
    {
        static void Main()
        {
            try
            {
                // Path to the encrypted XLSX file
                string encryptedFilePath = "encrypted.xlsx";

                // Verify the encrypted file exists
                if (!File.Exists(encryptedFilePath))
                {
                    Console.WriteLine($"File not found: {encryptedFilePath}");
                    return;
                }

                // Password used to encrypt the file
                string password = "myPassword";

                // LoadOptions with the password to open the encrypted workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };

                // Load the encrypted workbook using the password
                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Example verification: read the value of a cell that contains a formula
                Worksheet sheet = workbook.Worksheets[0];
                Cell formulaCell = sheet.Cells["A1"]; // assume A1 has a formula
                object calculatedValue = formulaCell.Value;

                // Output the calculated value for verification
                Console.WriteLine($"Calculated value in {formulaCell.Name}: {calculatedValue}");

                // Optionally, save the decrypted and recalculated workbook
                string decryptedFilePath = "decrypted.xlsx";
                workbook.Save(decryptedFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Decrypted workbook saved to {decryptedFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
