// Title: Verify that a workbook saved as an encrypted XLSX using Aspose.Cells for .NET is password‑protected by reopening it
// AI Prompts: Generate C# code that saves a Workbook with a password, then attempts to open the file without a password and returns a boolean indicating whether the file is encrypted. | Create a .NET snippet that uses Aspose.Cells LoadOptions to load an Excel file and catches CellsException to detect a missing password. | Write a method that accepts a file path, opens it with Aspose.Cells, and throws a custom exception if the workbook is not password‑protected.
// Common Searches: Aspose.Cells .NET check if saved XLSX is password protected after saving | C# detect encrypted Excel workbook by loading without password | How to programmatically confirm workbook encryption with Aspose.Cells | LoadOptions without password throws CellsException for protected file | Validate encryption status of an Excel file using Aspose.Cells API
// Tags: Aspose.Cells workbook encryption verification | C# load encrypted XLSX with LoadOptions | detect password‑protected Excel using CellsException | verify Excel file protection programmatically | Aspose.Cells password protection status check

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, assigns a password to encrypt it, saves it as an XLSX file, then attempts to reload the file without providing a password. If Aspose.Cells throws a CellsException, the code marks the workbook as encrypted and outputs the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set a password to encrypt the workbook when it is saved
            workbook.Settings.Password = "MySecretPassword";

            // Save the workbook as an encrypted XLSX file
            string encryptedFilePath = "encrypted.xlsx";

            try
            {
                workbook.Save(encryptedFilePath, SaveFormat.Xlsx);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                return;
            }

            // Verify the encryption status by attempting to load without a password
            bool isEncrypted = false;

            if (File.Exists(encryptedFilePath))
            {
                try
                {
                    // Attempt to load without providing a password
                    LoadOptions loadOptions = new LoadOptions(); // no password supplied
                    Workbook reopenedWorkbook = new Workbook(encryptedFilePath, loadOptions);
                    // If loading succeeds, the file is not password protected
                    isEncrypted = false;
                }
                catch (CellsException)
                {
                    // Loading failed because a password is required
                    isEncrypted = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"File not found: {encryptedFilePath}");
                return;
            }

            Console.WriteLine($"Workbook encrypted: {isEncrypted}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
