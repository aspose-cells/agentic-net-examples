// Title: Decrypt a Password‑Protected XLSX, Verify the Password, Recalculate Formulas, and Save with Aspose.Cells for .NET (C#)
// Description: This C# example shows how to use Aspose.Cells to (1) confirm a password for an encrypted XLSX file, (2) open the workbook with LoadOptions (Password and ParsingFormulaOnOpen enabled), (3) verify that the workbook is no longer encrypted, (4) recalculate all formulas, (5) read a formula cell to display its result, and (6) save the decrypted workbook to a new file.
// Keywords: Aspose.Cells decrypt XLSX | C# password protected Excel | verify Excel password Aspose.Cells | load workbook with password .NET | ParsingFormulaOnOpen | recalculate formulas Aspose.Cells | save decrypted workbook | FileFormatUtil.VerifyPassword | Excel encryption removal
// Common Searches: open password protected xlsx with Aspose.Cells C# | verify Excel file password using Aspose.Cells | recalculate formulas after decrypting Excel workbook | save decrypted Excel file with Aspose.Cells | how to use ParsingFormulaOnOpen in Aspose.Cells
// Developer Intent: Open an encrypted XLSX file, confirm the supplied password, recalculate its formulas, and write the decrypted workbook to disk.
// Use Cases: Validate a user‑provided password before loading a protected workbook to avoid runtime exceptions. | Ensure all formulas are up‑to‑date after decryption for accurate reporting or further processing. | Create an unencrypted copy of a secured Excel file for downstream systems that cannot handle encryption.
// AI Prompts: Write C# code that uses Aspose.Cells to verify a password, open an encrypted XLSX, recalculate all formulas, and save the decrypted workbook. | Explain how the ParsingFormulaOnOpen option influences formula evaluation when loading a password‑protected workbook. | Provide robust error‑handling patterns for loading and processing encrypted Excel files with Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example shows how to use Aspose.Cells to (1) confirm a password for an encrypted XLSX file, (2) open the workbook with LoadOptions (Password and ParsingFormulaOnOpen enabled), (3) verify that the workbook is no longer encrypted, (4) recalculate all formulas, (5) read a formula cell to display its result, and (6) save the decrypted workbook to a new file.
    class DecryptAndRecalculate
    {
        static void Main()
        {
            // Path to the encrypted XLSX file and its password
            string encryptedFilePath = "encrypted.xlsx";
            string password = "myPassword";

            // Ensure the input file exists
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"Error: File \"{encryptedFilePath}\" not found.");
                return;
            }

            try
            {
                // Verify that the supplied password is correct
                using (Stream stream = File.OpenRead(encryptedFilePath))
                {
                    bool isValid = FileFormatUtil.VerifyPassword(stream, password);
                    Console.WriteLine($"Password verification result: {isValid}");
                }

                // Load the workbook with the password; enable formula parsing on open
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password,
                    ParsingFormulaOnOpen = true
                };

                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                // Confirm that the workbook is no longer encrypted after loading
                Console.WriteLine($"Workbook IsEncrypted: {workbook.Settings.IsEncrypted}");

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Example: read a cell that contains a formula to verify recalculation
                Worksheet sheet = workbook.Worksheets[0];
                Cell formulaCell = sheet.Cells["A1"]; // adjust the address as needed
                Console.WriteLine($"Cell A1 formula: {formulaCell.Formula}");
                Console.WriteLine($"Cell A1 calculated value: {formulaCell.Value}");

                // Save the decrypted workbook (optional)
                string outputPath = "decrypted_output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Decrypted workbook saved to \"{outputPath}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
