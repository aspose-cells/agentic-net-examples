// Title: Decrypt a Password‑Protected XLSX and Recalculate Formulas with Aspose.Cells for .NET
// Description: C# sample that checks for an encrypted workbook, creates one if missing, verifies the supplied password using FileFormatUtil.VerifyPassword, opens the file with LoadOptions.Password, forces full formula recalculation, clears the protection flag, and saves an unencrypted copy while handling common exceptions.
// Keywords: Aspose.Cells decrypt XLSX | password protected Excel .NET | verify Excel password Aspose | calculate formulas Aspose.Cells | remove workbook password C# | load encrypted workbook Aspose | Excel encryption handling | FileFormatUtil.VerifyPassword | LoadOptions.Password example
// Common Searches: open password protected xlsx with Aspose.Cells | verify Excel file password programmatically | recalculate formulas after decrypting Excel | remove protection from Excel workbook using C# | Aspose.Cells sample for encrypted workbook
// Developer Intent: Open a protected XLSX, confirm the password, recalculate all formulas, strip the protection, and write an unencrypted file.
// Use Cases: Automated processing of secured Excel reports that require fresh calculations before distribution. | Batch conversion of encrypted workbooks to plain files while preserving calculated values. | Integration of password verification into a data‑pipeline that validates user‑provided Excel files.
// AI Prompts: Generate C# code to open an encrypted XLSX with a known password using Aspose.Cells, verify the password, recalculate all formulas, and save the workbook without protection. | Explain the role of FileFormatUtil.VerifyPassword and how LoadOptions.Password enables decryption in Aspose.Cells. | Suggest enhancements to keep cell styles, charts, and pivot tables intact while recalculating formulas after decryption.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that checks for an encrypted workbook, creates one if missing, verifies the supplied password using FileFormatUtil.VerifyPassword, opens the file with LoadOptions.Password, forces full formula recalculation, clears the protection flag, and saves an unencrypted copy while handling common exceptions.
class DecryptAndRecalculate
{
    static void Main()
    {
        // Paths for the encrypted and decrypted workbooks
        string encryptedFilePath = "encrypted.xlsx";
        string decryptedFilePath = "decrypted.xlsx";
        string password = "myPassword";

        try
        {
            // Ensure the encrypted workbook exists; if not, create a sample encrypted file
            if (!File.Exists(encryptedFilePath))
            {
                // Create a simple workbook with a formula
                Workbook sampleWb = new Workbook();
                Worksheet sheet = sampleWb.Worksheets[0];
                sheet.Cells["A1"].Formula = "=SUM(1, 2, 3)"; // result should be 6

                // Apply password protection
                sampleWb.Settings.Password = password;

                // Save the encrypted workbook
                sampleWb.Save(encryptedFilePath);
                Console.WriteLine($"Sample encrypted workbook created at: {encryptedFilePath}");
            }

            // Verify that the supplied password is correct
            bool isPasswordCorrect;
            using (Stream stream = File.OpenRead(encryptedFilePath))
            {
                isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
            }
            Console.WriteLine($"Password verification result: {isPasswordCorrect}");

            if (!isPasswordCorrect)
            {
                Console.WriteLine("The provided password is incorrect. Exiting.");
                return;
            }

            // Load the encrypted workbook with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Example: display a formula and its calculated value after recalculation
            Worksheet firstSheet = workbook.Worksheets[0];
            Cell formulaCell = firstSheet.Cells["A1"];
            Console.WriteLine($"Formula in A1: {formulaCell.Formula}");
            Console.WriteLine($"Calculated value in A1: {formulaCell.Value}");

            // Remove the encryption password to produce an unprotected copy
            workbook.Settings.Password = string.Empty;

            // Save the decrypted workbook
            workbook.Save(decryptedFilePath);
            Console.WriteLine($"Decrypted workbook saved to: {decryptedFilePath}");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (CellsException cellsEx)
        {
            Console.WriteLine($"Aspose.Cells error: {cellsEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
