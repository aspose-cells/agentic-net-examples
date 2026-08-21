// Title: Encrypt an Excel workbook with a user‑defined password using Aspose.Cells for .NET (AES‑128)
// Description: The sample prompts for a password, creates a new Workbook, writes a value to cell A1, sets the password, selects AES‑128 strong encryption, saves the file as EncryptedWorkbook.xlsx, and then reloads it with LoadOptions to verify the protection.
// Keywords: Aspose.Cells password protection | C# encrypt Excel file | AES 128 encryption Aspose | Workbook.Settings.Password | SetEncryptionOptions | LoadOptions password | protected .xlsx .NET | secure Excel workbook C# | North America developers | European .NET community
// Common Searches: How to add password protection to an Excel file with Aspose.Cells C# | Set AES‑128 encryption for a workbook in .NET | Open a password‑protected workbook using LoadOptions Aspose | C# code to encrypt and save Excel workbook | Aspose.Cells encrypt workbook example
// Developer Intent: Apply a user‑entered password and strong encryption to a newly created workbook and store it as a protected Excel file.
// Use Cases: Generate confidential reports that are automatically locked with a password before distribution. | Implement a secure export feature in a financial application using Aspose.Cells. | Validate encryption by programmatically reopening the saved file with the same password.
// AI Prompts: Provide C# code that encrypts a new Aspose.Cells workbook with a password and AES‑128, then saves it. | Show how to open a password‑protected Excel file using Aspose.Cells LoadOptions and read a specific cell. | Explain how to change the encryption algorithm or key size for an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// The sample prompts for a password, creates a new Workbook, writes a value to cell A1, sets the password, selects AES‑128 strong encryption, saves the file as EncryptedWorkbook.xlsx, and then reloads it with LoadOptions to verify the protection.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Prompt the user for a password to encrypt the workbook
        Console.Write("Enter password to encrypt workbook: ");
        string password = Console.ReadLine();

        // Create a new workbook (lifecycle: create)
        Workbook wb = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Encrypted content");

        // Set the password that will be required to open the workbook
        wb.Settings.Password = password;

        // Optionally specify stronger encryption (AES 128-bit)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to a file (lifecycle: save)
        string outputPath = "EncryptedWorkbook.xlsx";
        wb.Save(outputPath, SaveFormat.Xlsx);

        // Verify that the workbook is encrypted by loading it with the password (lifecycle: load)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedWb = new Workbook(outputPath, loadOptions);
        Console.WriteLine("Loaded cell value: " + loadedWb.Worksheets[0].Cells["A1"].StringValue);
    }
}
