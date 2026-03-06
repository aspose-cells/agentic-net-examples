using System;
using Aspose.Cells;

class DecryptExcel
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string encryptedFile = "protected.xlsx";

        // Path where the decrypted file will be saved
        string decryptedFile = "unprotected.xlsx";

        // LoadOptions with the password required to open the encrypted workbook
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "myPassword";

        // Open the workbook using the load options
        Workbook workbook = new Workbook(encryptedFile, loadOptions);

        // Remove the opening password so the workbook will be saved unencrypted
        workbook.Settings.Password = null;

        // Save the workbook without password protection
        workbook.Save(decryptedFile);
    }
}