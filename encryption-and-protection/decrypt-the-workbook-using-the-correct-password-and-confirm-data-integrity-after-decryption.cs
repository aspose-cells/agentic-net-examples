// Title: Decrypt a password‑protected Excel workbook with Aspose.Cells for .NET and verify data integrity
// Description: This example demonstrates how to open an encrypted Excel file (encrypted.xlsx) using Aspose.Cells LoadOptions with the correct password, confirm that the workbook is no longer encrypted via workbook.Settings.IsEncrypted, read a known cell (A1) to ensure the data is intact, remove the password, and save the workbook as an unprotected file (decrypted_unprotected.xlsx).
// Keywords: Aspose.Cells | C# | .NET | decrypt Excel workbook | password protected Excel | LoadOptions.Password | verify cell value | remove workbook encryption | save unencrypted Excel | workbook.Settings.IsEncrypted
// Common Searches: Aspose.Cells open password protected Excel file C# | How to remove password from Excel workbook using Aspose.Cells | Check if workbook is encrypted after loading Aspose.Cells | Read cell value after decrypting Excel with Aspose | Save decrypted Excel without password Aspose.Cells
// Developer Intent: Load an encrypted workbook with the proper password, confirm decryption, validate cell data, and optionally save a non‑encrypted copy.
// Use Cases: Load "encrypted.xlsx" using LoadOptions.Password = "mySecret" and verify workbook.Settings.IsEncrypted returns false. | Read Worksheets[0].Cells["A1"].StringValue after decryption to confirm expected content. | Clear workbook.Settings.Password and save as "decrypted_unprotected.xlsx" to produce an unencrypted file.
// AI Prompts: Write C# code that opens a password‑protected Excel file with Aspose.Cells, checks the encryption flag, reads a specific cell for verification, and saves the workbook without a password. | Explain error handling strategies for invalid passwords or corrupted files when loading encrypted workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to open an encrypted Excel file (encrypted.xlsx) using Aspose.Cells LoadOptions with the correct password, confirm that the workbook is no longer encrypted via workbook.Settings.IsEncrypted, read a known cell (A1) to ensure the data is intact, remove the password, and save the workbook as an unprotected file (decrypted_unprotected.xlsx).
class DecryptWorkbookDemo
{
    static void Main()
    {
        // Path to the encrypted workbook file
        string encryptedFilePath = "encrypted.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"Error: File \"{encryptedFilePath}\" not found.");
            return;
        }

        // The password that was used to encrypt the workbook
        string password = "mySecret";

        try
        {
            // Load the encrypted workbook by providing the password in LoadOptions
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // After loading with the correct password, the workbook should be decrypted
            Console.WriteLine("Is workbook encrypted after load? " + workbook.Settings.IsEncrypted);

            // Confirm data integrity by reading a known cell value (e.g., A1)
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Value of cell A1 after decryption: " + cellValue);

            // Optional: remove the encryption password and save an unprotected copy
            // Clearing the password ensures the workbook is saved without encryption
            workbook.Settings.Password = null;
            workbook.Save("decrypted_unprotected.xlsx");
            Console.WriteLine("Decrypted workbook saved as \"decrypted_unprotected.xlsx\".");
        }
        catch (Exception ex)
        {
            // Handle any errors (e.g., invalid password, file format issues)
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
