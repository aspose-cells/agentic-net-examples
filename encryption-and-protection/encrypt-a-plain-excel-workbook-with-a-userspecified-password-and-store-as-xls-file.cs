using System;
using Aspose.Cells;

class EncryptWorkbookToXls
{
    static void Main()
    {
        // Prompt user for the password to protect the workbook
        Console.Write("Enter password to encrypt the workbook: ");
        string password = Console.ReadLine();

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive information");

        // Set the password that will be required to open the workbook
        workbook.Settings.Password = password;

        // Apply stronger encryption options for the Excel 97‑2003 format
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook as an XLS file (Excel 97‑2003)
        workbook.Save("EncryptedWorkbook.xls", SaveFormat.Excel97To2003);

        // Verify the encryption by loading the file with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedWorkbook = new Workbook("EncryptedWorkbook.xls", loadOptions);
        Console.WriteLine("Loaded cell value: " + loadedWorkbook.Worksheets[0].Cells["A1"].Value);
    }
}