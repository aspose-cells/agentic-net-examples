using System;
using Aspose.Cells;

namespace WorkbookEncryptionLogger
{
    class Program
    {
        // Simple logger that writes to console with timestamp
        static void LogEncryptionStatus(string action, bool isEncrypted)
        {
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {action}: IsEncrypted = {isEncrypted}");
        }

        static void Main(string[] args)
        {
            // Path for the workbook file
            string filePath = "EncryptedWorkbook.xlsx";

            // -------------------------------------------------
            // 1. Create a new workbook (create rule)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Initial encryption status (should be false)
            LogEncryptionStatus("Initial check", workbook.Settings.IsEncrypted);

            // -------------------------------------------------
            // 2. Enable encryption by setting a password
            // -------------------------------------------------
            workbook.Settings.Password = "MySecretPassword";

            // Save the workbook (save rule)
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Log after setting password
            LogEncryptionStatus("After setting password", workbook.Settings.IsEncrypted);

            // -------------------------------------------------
            // 3. Load the encrypted workbook (load rule)
            // -------------------------------------------------
            LoadOptions loadOptions = new LoadOptions { Password = "MySecretPassword" };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Verify encryption status after load
            LogEncryptionStatus("After loading encrypted workbook", loadedWorkbook.Settings.IsEncrypted);

            // -------------------------------------------------
            // 4. Remove encryption by clearing the password
            // -------------------------------------------------
            loadedWorkbook.Settings.Password = null; // Clearing password removes encryption

            // Save the unencrypted workbook (save rule)
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            loadedWorkbook.Save(unencryptedPath, SaveFormat.Xlsx);

            // Log after removing password
            LogEncryptionStatus("After removing password", loadedWorkbook.Settings.IsEncrypted);

            // -------------------------------------------------
            // 5. Load the unencrypted workbook to confirm status
            // -------------------------------------------------
            Workbook finalWorkbook = new Workbook(unencryptedPath);
            LogEncryptionStatus("After loading unencrypted workbook", finalWorkbook.Settings.IsEncrypted);
        }
    }
}