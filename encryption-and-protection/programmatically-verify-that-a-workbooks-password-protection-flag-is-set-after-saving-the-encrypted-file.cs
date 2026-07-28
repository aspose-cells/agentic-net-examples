// Title: Check Excel workbook encryption flag after saving with Aspose.Cells for .NET
// Description: Creates a workbook, applies a password via Workbook.Settings.Password, saves it, then confirms encryption using Workbook.Settings.IsEncrypted, LoadOptions, and FileFormatUtil.DetectFileFormat without reopening the file.
// Keywords: Aspose.Cells encryption | C# verify Excel password protection | Workbook.Settings.IsEncrypted | FileFormatUtil DetectFileFormat | LoadOptions password Aspose.Cells | Excel file encryption flag .NET | detect encrypted .xlsx Aspose
// Common Searches: how to verify Excel file is encrypted with Aspose.Cells | Aspose.Cells check workbook password protection after save | detect encrypted workbook without opening in C# | Workbook.Settings.IsEncrypted returns true | load encrypted Excel using LoadOptions password
// Developer Intent: Confirm that a workbook saved with a password has its encryption flag set to true.
// Use Cases: Open the saved file with the correct password via LoadOptions and read Workbook.Settings.IsEncrypted. | Use FileFormatUtil.DetectFileFormat to obtain FileFormatInfo.IsEncrypted, allowing validation without loading the workbook. | Attempt to load the file without a password, catch CellsException, and ensure access is denied.
// AI Prompts: Generate C# code that saves an Excel workbook with a password using Aspose.Cells and then verifies the encryption flag with Workbook.Settings.IsEncrypted and FileFormatInfo.IsEncrypted. | Explain how to handle CellsException when trying to open a password‑protected workbook without providing a password in Aspose.Cells. | Show how to use LoadOptions to open an encrypted workbook and programmatically confirm its encrypted status.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, applies a password via Workbook.Settings.Password, saves it, then confirms encryption using Workbook.Settings.IsEncrypted, LoadOptions, and FileFormatUtil.DetectFileFormat without reopening the file.
class VerifyWorkbookEncryption
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted content");

            // Set the encryption password
            workbook.Settings.Password = "mySecretPassword";

            // Save the encrypted workbook
            string encryptedFilePath = "encryptedWorkbook.xlsx";
            workbook.Save(encryptedFilePath);

            // Verify that the file exists before attempting to load it
            if (!File.Exists(encryptedFilePath))
            {
                Console.WriteLine($"File not found: {encryptedFilePath}");
                return;
            }

            // Attempt to load the encrypted workbook without a password
            try
            {
                Workbook loadedWithoutPassword = new Workbook(encryptedFilePath);
                Console.WriteLine("Workbook loaded without password (unexpected).");
                Console.WriteLine("Workbook.Settings.IsEncrypted: " + loadedWithoutPassword.Settings.IsEncrypted);
            }
            catch (CellsException)
            {
                // Expected exception for encrypted file without password
                Console.WriteLine("Cannot load encrypted workbook without password (as expected).");
            }

            // Load the workbook with the correct password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "mySecretPassword"
            };
            Workbook loadedWithPassword = new Workbook(encryptedFilePath, loadOptions);
            Console.WriteLine("Workbook loaded with password successfully.");
            Console.WriteLine("Workbook.Settings.IsEncrypted: " + loadedWithPassword.Settings.IsEncrypted);

            // Use FileFormatInfo to check encryption status without loading the workbook
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedFilePath);
            Console.WriteLine("FileFormatInfo.IsEncrypted: " + formatInfo.IsEncrypted);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
