// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and confirm that opening it without the password fails with CellsException
// AI Prompts: Write a C# program that creates a Workbook, assigns a password via workbook.Settings.Password, saves it as an encrypted .xlsx, then tries to load the file without a password using LoadOptions and catches the resulting CellsException. | Adapt the code to encrypt the workbook as an .xls file instead of .xlsx and verify that loading it without the password still triggers a CellsException. | Add detailed logging to the example so that when the password‑protected workbook is opened without credentials, the exception type and message are printed to the console.
// Common Searches: how to encrypt an Excel file with a password using Aspose.Cells C# | Aspose.Cells throws CellsException when opening password protected workbook without password | C# example to save encrypted .xlsx and test opening without credentials | catch CellsException for protected Excel files in .NET | verify that encrypted workbook cannot be opened without password Aspose.Cells
// Tags: Aspose.Cells encrypt workbook with password | C# load password‑protected XLSX with LoadOptions | CellsException handling for encrypted Excel | verify workbook encryption failure .NET | Aspose.Cells password protection error handling

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, adds data, sets a password via workbook.Settings.Password, saves it as an encrypted Encrypted.xlsx, then attempts to open the file without providing the password using LoadOptions. The operation throws a CellsException, which is caught and logged to demonstrate that decryption without the correct password fails.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Path for the encrypted workbook
        string filePath = "Encrypted.xlsx";

        try
        {
            // -------------------------------------------------
            // 1. Create a new workbook using Aspose.Cells
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // -------------------------------------------------
            // 2. Encrypt the workbook with a password
            // -------------------------------------------------
            // Set the password that will be required to open the file
            workbook.Settings.Password = "Secret123";

            // Save the encrypted workbook
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at '{Path.GetFullPath(filePath)}'.");

            // -------------------------------------------------
            // 3. Attempt to open the encrypted file without providing the password – this should fail
            // -------------------------------------------------
            if (File.Exists(filePath))
            {
                try
                {
                    // LoadOptions without password – expected to throw
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                    Workbook protectedWb = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Unexpectedly succeeded in opening the encrypted file without a password.");
                }
                catch (CellsException ex)
                {
                    // Aspose.Cells throws CellsException when a password is required
                    Console.WriteLine("Failed to open encrypted workbook without password as expected.");
                    Console.WriteLine($"Exception type: {ex.GetType().Name}");
                    Console.WriteLine($"Message: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An unexpected error occurred while opening the encrypted workbook.");
                    Console.WriteLine($"Exception type: {ex.GetType().Name}");
                    Console.WriteLine($"Message: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
        catch (Exception ex)
        {
            // General exception handling for any unexpected errors
            Console.WriteLine("An error occurred during the workbook encryption demo.");
            Console.WriteLine($"Exception type: {ex.GetType().Name}");
            Console.WriteLine($"Message: {ex.Message}");
        }
    }
}
