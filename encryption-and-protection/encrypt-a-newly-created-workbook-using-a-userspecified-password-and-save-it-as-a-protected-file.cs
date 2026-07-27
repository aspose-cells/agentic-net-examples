// Title: Encrypt an Aspose.Cells Workbook with a User‑Provided Password in C#
// Description: C# example that prompts for a password, creates a new Workbook, writes sample data, applies strong 128‑bit encryption via SetEncryptionOptions, saves as a protected XLSX file, and verifies the protection by reopening the file with LoadOptions.Password.
// Keywords: Aspose.Cells encrypt workbook | C# password protect Excel | SetEncryptionOptions Aspose.Cells | strong cryptographic provider | load password protected workbook | SaveFormat.Xlsx encryption | Workbook.Settings.Password
// Common Searches: How to password protect an Excel file using Aspose.Cells .NET | Aspose.Cells set strong encryption for XLSX | Load a password‑protected workbook with Aspose.Cells C# | Change encryption algorithm and key size in Aspose.Cells | Verify encrypted workbook Aspose.Cells example
// Developer Intent: Create a new workbook, encrypt it with a user‑supplied password, and save it as a protected Excel file.
// Use Cases: Prompt the end‑user for a password, assign it to wb.Settings.Password, and save the workbook as an encrypted .xlsx file. | Apply 128‑bit StrongCryptographicProvider encryption for compatibility with older Excel versions using wb.SetEncryptionOptions. | Reload the saved file with LoadOptions.Password to confirm that the encryption works and cell data is accessible.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, asks the user for a password, applies 128‑bit strong encryption, and saves it as a protected XLSX file. | Show how to open a password‑protected Excel workbook with Aspose.Cells in C# and handle incorrect password exceptions gracefully. | Explain how to switch between different encryption types and key sizes when protecting a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that prompts for a password, creates a new Workbook, writes sample data, applies strong 128‑bit encryption via SetEncryptionOptions, saves as a protected XLSX file, and verifies the protection by reopening the file with LoadOptions.Password.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Prompt the user for a password to encrypt the workbook
        Console.Write("Enter password to encrypt workbook: ");
        string password = Console.ReadLine();

        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Encrypted content");

        // Set the password that will be required to open the workbook
        wb.Settings.Password = password;

        // (Optional) Specify stronger encryption options for older Excel formats
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook as an encrypted file
        string outputPath = "EncryptedWorkbook.xlsx";
        wb.Save(outputPath, SaveFormat.Xlsx);

        // Verify encryption by loading the workbook with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook loadedWb = new Workbook(outputPath, loadOptions);
        Console.WriteLine("Loaded cell value: " + loadedWb.Worksheets[0].Cells["A1"].StringValue);
    }
}
