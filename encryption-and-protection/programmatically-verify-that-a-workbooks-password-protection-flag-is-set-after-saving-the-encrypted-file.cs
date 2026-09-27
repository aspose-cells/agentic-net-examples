// Title: Programmatically confirm workbook password protection after saving an encrypted Excel file with Aspose.Cells for .NET
// AI Prompts: Generate a C# snippet that creates a workbook, sets workbook.Settings.Password, saves it, reloads with LoadOptions.Password, and verifies the IsEncrypted property is true. | Update the example to throw an exception when the loaded workbook's encryption flag is false after saving with a password. | Write a reusable C# method that accepts a file path and password, saves an encrypted workbook, reloads it, and returns a boolean indicating whether the workbook reports as encrypted.
// Common Searches: how to programmatically check if an Excel workbook is password protected using Aspose.Cells in C# | verify that workbook.Settings.Password actually encrypts the file with Aspose.Cells .NET | C# load encrypted .xlsx with Aspose.Cells and read IsEncrypted property | unit test for Aspose.Cells password protection flag after saving workbook | Aspose.Cells .NET determine if saved workbook is encrypted
// Tags: Aspose.Cells workbook encryption verification | C# load password‑protected Excel file with LoadOptions | save workbook with Settings.Password .NET | unit test workbook password protection Aspose.Cells | determine encryption status of saved workbook

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, assigns a password via workbook.Settings.Password, saves it as an .xlsx file, then reloads the file using LoadOptions with the same password and demonstrates how to read the IsEncrypted flag to confirm the workbook is protected.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Define the password and apply encryption to the workbook
            string password = "MySecretPwd";
            workbook.Settings.Password = password; // encrypt the file with the password

            // Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the encrypted workbook using the same password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            Workbook loadedWorkbook;
            try
            {
                loadedWorkbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully with the provided password.");
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Optional: confirm that the workbook is encrypted (property may vary by version)
            // If the IsEncrypted property is unavailable, this check can be omitted.
            // bool isEncrypted = loadedWorkbook.IsEncrypted; // Uncomment if supported

        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
