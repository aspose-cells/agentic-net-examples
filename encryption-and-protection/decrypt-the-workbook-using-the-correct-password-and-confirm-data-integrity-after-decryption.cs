// Title: Decrypt a password‑protected XLSX workbook and validate its contents using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load encrypted XLSX with password and read cell A1 | how to remove password protection from an Excel file using Aspose.Cells .NET | verify workbook data after decryption with Aspose.Cells in C# | calculate checksum of numeric cells after opening a protected Excel workbook | save decrypted workbook without password using Aspose.Cells
// Tags: password protected workbook loading Aspose.Cells | excel decryption C# Aspose.Cells | post‑decryption cell verification Aspose.Cells | numeric cell checksum calculation Aspose.Cells | saving workbook without password Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to open an encrypted XLSX file with a password using Aspose.Cells LoadOptions, read a known cell to confirm successful decryption, optionally compute a checksum of all numeric cells in the first worksheet, and then save the workbook as an unprotected file.
class Program
{
    static void Main()
    {
        // Paths and password
        string encryptedPath = "encrypted.xlsx";
        string decryptedPath = "decrypted.xlsx";
        string password = "myPassword";

        // Verify the encrypted file exists
        if (!File.Exists(encryptedPath))
        {
            Console.WriteLine($"Error: File \"{encryptedPath}\" not found.");
            return;
        }

        try
        {
            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedPath, loadOptions);

            // Confirm data integrity: read a known cell value
            Worksheet sheet = workbook.Worksheets[0];
            string cellValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine($"Cell A1 value after decryption: {cellValue}");

            // Optional: compute a simple checksum of all numeric cells
            double checksum = 0;
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Type == CellValueType.IsNumeric)
                    checksum += cell.DoubleValue;
            }
            Console.WriteLine($"Checksum of numeric values: {checksum}");

            // Save the workbook without a password to verify it is decrypted
            workbook.Save(decryptedPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved decrypted successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
