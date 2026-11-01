// Title: Verify that a saved XLSX workbook is password‑encrypted by reopening it with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to confirm an XLSX file is password protected after saving | detect encryption of an Excel workbook when loading without a password using Aspose.Cells | Aspose.Cells .NET check if saved workbook requires a password to open | validate workbook encryption status programmatically with Aspose.Cells | catch CellsException for encrypted Excel file in C# Aspose.Cells example
// Tags: verify workbook encryption Aspose.Cells | load encrypted XLSX without password | catch CellsException password protection | check encryption status after save Aspose.Cells | password protected Excel file validation C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds data, sets a password to encrypt it, saves the file as XLSX, then attempts to reload the file without providing a password. It catches the expected CellsException that indicates the workbook is encrypted and outputs the encryption status, demonstrating how to programmatically verify password protection of a saved Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel file
            string filePath = "EncryptedWorkbook.xlsx";

            // -----------------------------------------------------------------
            // Create a new workbook and encrypt it with a password
            // -----------------------------------------------------------------
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // Set password to encrypt the workbook
            workbook.Settings.Password = "MySecretPassword";

            // Save the workbook (encryption will be applied automatically)
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file was created
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The workbook file was not created.", filePath);

            // -----------------------------------------------------------------
            // Reopen the saved file without providing a password
            // and verify that it is encrypted
            // -----------------------------------------------------------------
            bool isEncrypted = false;

            try
            {
                // Attempt to load without a password; this should fail if encrypted
                Workbook loadedWorkbook = new Workbook(filePath);
                // If loading succeeds, the workbook is not encrypted
                isEncrypted = false;
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                // Expected exception for encrypted workbook
                isEncrypted = true;
            }
            catch (Exception ex)
            {
                // Unexpected exception
                Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
                throw;
            }

            // Output the result
            Console.WriteLine($"Is the workbook encrypted? {isEncrypted}");

            // Optional: assert the expected condition
            if (!isEncrypted)
                throw new InvalidOperationException("The workbook was expected to be encrypted but is not.");
        }
        catch (Exception ex)
        {
            // General error handling
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
