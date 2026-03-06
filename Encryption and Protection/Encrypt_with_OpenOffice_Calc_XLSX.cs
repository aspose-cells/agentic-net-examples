using System;
using Aspose.Cells;

class EncryptWorkbook
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some sample data
        wb.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

        // Set a password to encrypt the workbook (OpenOffice Calc will prompt for this password)
        wb.Settings.Password = "OpenOfficePwd";

        // Set encryption options (for XLSX this uses the default SHA‑AES encryption; the enum value is ignored)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook as XLSX
        wb.Save("EncryptedWorkbook.xlsx");

        // Load the encrypted workbook using the password to verify it works
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "OpenOfficePwd";
        Workbook loadedWb = new Workbook("EncryptedWorkbook.xlsx", loadOptions);

        // Output the value to confirm successful decryption
        Console.WriteLine("Cell A1 value: " + loadedWb.Worksheets[0].Cells["A1"].StringValue);
    }
}