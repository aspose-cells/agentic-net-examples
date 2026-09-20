// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and verify that opening it without a password raises a CellsException
// AI Prompts: Create a Workbook, assign a password via Workbook.Settings.Password, and save it as an encrypted .xlsx file with Aspose.Cells. | Load the encrypted .xlsx file without providing a password and capture the CellsException to confirm access is denied. | Change the password value, re‑save the workbook, and repeat the load test to ensure the exception behavior remains consistent.
// Common Searches: Aspose.Cells .NET how to set a password for an Excel workbook and save it encrypted | C# catch CellsException when opening a password‑protected Excel file with Aspose.Cells | verify that an encrypted workbook cannot be opened without a password using Aspose.Cells | example code for workbook.Settings.Password encryption Aspose.Cells | test workbook encryption access denial Aspose.Cells C#
// Tags: Workbook.Settings.Password encryption Aspose.Cells | save encrypted Xlsx with Aspose.Cells | load password‑protected workbook without password | handle CellsException for encrypted Excel file | verify workbook access denial Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, applies a password using Workbook.Settings.Password, saves it as an encrypted XLSX file, then attempts to open the file without a password, catching the expected CellsException to demonstrate that access is denied.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // 2. Set password protection for the workbook
            //    Using WorkbookSettings.Password works across Aspose.Cells versions
            workbook.Settings.Password = "SecretPassword123";

            // 3. Save the workbook with encryption
            string encryptedFilePath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at: {encryptedFilePath}");

            // 4. Attempt to open the encrypted workbook without providing a password
            try
            {
                if (File.Exists(encryptedFilePath))
                {
                    // This should fail because no password is supplied
                    Workbook openedWithoutPassword = new Workbook(encryptedFilePath);
                    Console.WriteLine("Unexpectedly opened the encrypted workbook without a password.");
                }
                else
                {
                    Console.WriteLine($"File not found: {encryptedFilePath}");
                }
            }
            catch (CellsException ex)
            {
                // Expected outcome: access denied due to missing password
                Console.WriteLine("Access denied as expected. Exception message:");
                Console.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
        catch (Exception ex)
        {
            // General exception handling for any unexpected errors during creation or saving
            Console.WriteLine("An error occurred during workbook processing:");
            Console.WriteLine(ex.Message);
        }
    }
}
