using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Prompt the user for a password to protect the workbook
        Console.Write("Enter password to protect the workbook: ");
        string password = Console.ReadLine();

        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("This workbook is encrypted.");

        // Set the password that will be required to open the workbook
        workbook.Settings.Password = password;

        // Optionally specify stronger encryption (e.g., 128‑bit AES)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to a file (lifecycle rule: save)
        string outputPath = "ProtectedWorkbook.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Load the protected workbook using the password to verify (lifecycle rule: load)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);

        // Output a cell value to confirm successful decryption
        Console.WriteLine("Loaded cell value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}