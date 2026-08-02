// Title: Decrypt a Password‑Protected XLSX and Recalculate Formulas with Aspose.Cells for .NET
// Description: C# sample that verifies a password for an encrypted XLSX, opens the workbook with parsing formulas on load, recalculates all formulas, displays A1's formula and value, and saves an unprotected copy.
// Keywords: Aspose.Cells decrypt XLSX | C# password protected Excel | verify Excel password Aspose | load workbook with password | ParsingFormulaOnOpen | recalculate formulas Aspose.Cells | save decrypted workbook .NET
// Common Searches: open encrypted Excel file with Aspose.Cells C# | check password before loading workbook Aspose | recalculate formulas after opening protected XLSX | save decrypted copy of password‑protected workbook | Aspose.Cells ParsingFormulaOnOpen example
// Developer Intent: Load a password‑protected XLSX, confirm the password, recalculate all formulas, and write an unencrypted version.
// Use Cases: Validate user‑supplied passwords to avoid runtime errors when processing secured workbooks. | Ensure that all formulas reflect current data after decryption for reporting or analytics. | Create a plain‑text copy of a protected workbook for downstream automation or archival.
// AI Prompts: Write C# code using Aspose.Cells to open an encrypted XLSX with a known password, verify the password, recalculate formulas, and save a decrypted file. | Explain how to handle an invalid password when loading a protected workbook with Aspose.Cells and log the result. | Describe the effect of LoadOptions.ParsingFormulaOnOpen on formula evaluation in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that verifies a password for an encrypted XLSX, opens the workbook with parsing formulas on load, recalculates all formulas, displays A1's formula and value, and saves an unprotected copy.
class DecryptAndRecalculate
{
    static void Main()
    {
        try
        {
            // Path to the encrypted XLSX file and its password
            string encryptedPath = "encrypted.xlsx";
            string password = "myPassword";

            // Ensure the encrypted file exists
            if (!File.Exists(encryptedPath))
            {
                Console.WriteLine($"File not found: {encryptedPath}");
                return;
            }

            // Verify that the provided password is correct
            bool isValid;
            using (Stream stream = File.OpenRead(encryptedPath))
            {
                isValid = FileFormatUtil.VerifyPassword(stream, password);
            }
            Console.WriteLine($"Password verification result: {isValid}");

            if (!isValid)
            {
                Console.WriteLine("Invalid password. Cannot open the workbook.");
                return;
            }

            // Load the workbook with the password; formulas will be parsed on open
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                Password = password,
                ParsingFormulaOnOpen = true
            };

            Workbook workbook = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine($"Workbook loaded. IsEncrypted: {workbook.Settings.IsEncrypted}");

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Example: display formula and value of cell A1 after recalculation
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine($"A1 formula: {cell.Formula}");
            Console.WriteLine($"A1 value after recalculation: {cell.Value}");

            // Save the decrypted workbook to a new file
            string decryptedPath = "decrypted.xlsx";
            workbook.Save(decryptedPath);
            Console.WriteLine($"Decrypted workbook saved to {decryptedPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
