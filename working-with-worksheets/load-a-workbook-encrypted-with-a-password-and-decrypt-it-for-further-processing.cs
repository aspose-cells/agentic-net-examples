using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookDecryptionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the encrypted workbook file
                string encryptedFilePath = "encrypted.xlsx";

                // Password used to encrypt the workbook
                string password = "mySecretPassword";

                // Verify that the input file exists
                if (!File.Exists(encryptedFilePath))
                {
                    Console.WriteLine($"Input file not found: {encryptedFilePath}");
                    return;
                }

                // LoadOptions with the password for opening the protected workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                loadOptions.Password = password;

                // Load the password‑protected workbook
                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                // Remove the encryption password so the workbook can be saved unprotected
                workbook.Settings.Password = null;

                // Example processing: read a cell value
                Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);

                // Save the workbook without password protection
                string unprotectedFilePath = "unprotected.xlsx";
                workbook.Save(unprotectedFilePath);
                Console.WriteLine($"Workbook saved without password: {unprotectedFilePath}");
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