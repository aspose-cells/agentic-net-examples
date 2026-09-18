// Title: Decrypt a password‑protected XLSX workbook, recalculate all formulas, and save the result using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an encrypted .xlsx file with a known password via Aspose.Cells LoadOptions, forces full workbook formula calculation, reads a specific cell value, and saves the decrypted workbook. | Generate a C# example that loads a password‑protected Excel file, triggers CalculateFormula, verifies the output of a target cell, and writes the unprotected file to disk using Aspose.Cells.
// Common Searches: asp.net how to open an encrypted xlsx file with a password and recalculate formulas using Aspose.Cells | c# decrypt password protected Excel workbook and force formula evaluation with Aspose.Cells | example code to load encrypted XLSX, calculate all formulas, and save decrypted copy in C#
// Tags: load encrypted xlsx with LoadOptions Aspose.Cells | calculate workbook formulas Aspose.Cells | save decrypted workbook as xlsx Aspose.Cells | c# password protected excel decryption Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for an encrypted XLSX file, loads it using LoadOptions with the supplied password, forces calculation of all formulas, reads the value of cell A1 for verification, and saves the decrypted and recalculated workbook to a new file.
class DecryptAndRecalculate
{
    static void Main()
    {
        try
        {
            // Path to the encrypted XLSX file
            string encryptedFilePath = "encrypted.xlsx";

            // Verify that the encrypted file exists
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"Error: The file \"{encryptedFilePath}\" was not found.");
                return;
            }

            // Known password for the encrypted file
            string password = "YourPassword";

            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Force calculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Example verification: read the value of a specific cell after recalculation
            Worksheet sheet = workbook.Worksheets[0]; // first worksheet
            Cell targetCell = sheet.Cells["A1"]; // cell to verify
            Console.WriteLine($"Value of {targetCell.Name} after recalculation: {targetCell.Value}");

            // Optionally, save the decrypted and recalculated workbook to a new file
            string decryptedFilePath = "decrypted.xlsx";
            workbook.Save(decryptedFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Decrypted workbook saved to \"{decryptedFilePath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
