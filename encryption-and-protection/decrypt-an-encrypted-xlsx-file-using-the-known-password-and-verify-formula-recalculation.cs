// Title: Decrypt a password‑protected XLSX workbook and force full formula recalculation with Aspose.Cells for .NET
// Common Searches: how to open a password protected xlsx file using Aspose.Cells in C# | decrypt Excel workbook with known password and recalculate formulas Aspose.Cells | Aspose.Cells C# load encrypted workbook and save as unencrypted file | handle CellsException for incorrect password when loading Excel in .NET | force workbook.CalculateFormula after opening encrypted Excel with Aspose.Cells
// Tags: Aspose.Cells LoadOptions password decryption | Workbook.CalculateFormula after decryption | save workbook as unencrypted XLSX C# | CellsException handling for password errors | open protected Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the existence of an encrypted XLSX file, opens it with Aspose.Cells using LoadOptions.Password, forces a full workbook formula recalculation, reads the recalculated value of a sample cell, and optionally saves the decrypted workbook as a new XLSX file while gracefully handling password‑related CellsException errors.
class DecryptAndRecalculate
{
    static void Main()
    {
        // Path to the encrypted XLSX file
        string encryptedFilePath = "encrypted.xlsx";

        // Password used to encrypt the file
        string password = "YourPassword";

        // Verify that the input file exists
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"File not found: {encryptedFilePath}");
            return;
        }

        try
        {
            // LoadOptions with the password to open the encrypted workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the encrypted workbook using the password
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Force formula recalculation for the entire workbook
            workbook.CalculateFormula();

            // Example verification: read a cell that contains a formula
            Worksheet sheet = workbook.Worksheets[0];
            Cell formulaCell = sheet.Cells["A1"]; // assume A1 has a formula
            Console.WriteLine($"Recalculated value in {formulaCell.Name}: {formulaCell.Value}");

            // Save the decrypted workbook (optional)
            string decryptedFilePath = "decrypted.xlsx";
            workbook.Save(decryptedFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Decrypted workbook saved to {decryptedFilePath}");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for password issues; check the message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Invalid password provided for the encrypted workbook.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
