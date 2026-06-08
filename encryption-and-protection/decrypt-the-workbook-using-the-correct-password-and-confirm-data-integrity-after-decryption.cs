using System;
using System.IO;
using Aspose.Cells;

class DecryptWorkbookDemo
{
    static void Main()
    {
        // Paths
        string encryptedPath = "encrypted.xlsx";
        string unprotectedPath = "unprotected.xlsx";
        string password = "mySecret";

        // Verify the encrypted file exists
        if (!File.Exists(encryptedPath))
        {
            Console.WriteLine($"File not found: {encryptedPath}");
            return;
        }

        try
        {
            // Load the workbook with the supplied password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook workbook = new Workbook(encryptedPath, loadOptions);

            // Confirm encryption status after load
            Console.WriteLine("Workbook IsEncrypted (after load): " + workbook.Settings.IsEncrypted);

            // Read a cell value to verify data integrity after decryption
            string originalValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Cell A1 value: " + originalValue);

            // Remove encryption by clearing the password
            workbook.Settings.Password = null;

            // Save the unprotected workbook
            workbook.Save(unprotectedPath);
            Console.WriteLine($"Workbook saved without protection to: {unprotectedPath}");

            // Reload the saved workbook to ensure it is no longer encrypted
            Workbook unprotectedWorkbook = new Workbook(unprotectedPath);
            Console.WriteLine("Workbook IsEncrypted (after save): " + unprotectedWorkbook.Settings.IsEncrypted);

            // Verify that the cell value remains unchanged
            string reloadedValue = unprotectedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Data integrity check passed: " + (originalValue == reloadedValue));
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for invalid password or other workbook errors
            Console.WriteLine($"CellsException: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}