// Title: C# – Encrypt an ODS workbook with a password and confirm protection using Aspose.Cells
// Description: This example creates a Workbook, writes sample data, applies a password via Workbook.Settings.Password, saves the file as ODS, checks the IsEncrypted flag with FileFormatUtil, and then reloads the file using LoadOptions.Password to demonstrate that the workbook opens only after the correct password is supplied.
// Keywords: Aspose.Cells | .NET | C# | ODS encryption | password‑protected ODS | Workbook.Settings.Password | FileFormatUtil.IsEncrypted | LoadOptions.Password | secure spreadsheet export
// Common Searches: Aspose.Cells encrypt ODS file C# | detect encrypted ODS workbook Aspose | load password protected ODS with Aspose.Cells | set password for ODS using Aspose.Cells .NET | verify ODS file encryption programmatically
// Developer Intent: Add password protection to an ODS workbook and ensure it requires the password when opened.
// Use Cases: Distribute confidential spreadsheets in ODS format with built‑in encryption | Automate compliance checks by confirming the IsEncrypted flag after saving | Read a protected ODS file in a downstream application after providing the correct password
// AI Prompts: Generate C# code that encrypts an ODS workbook with a password using Aspose.Cells and verifies the encryption status. | Show how to open a password‑protected ODS file with Aspose.Cells, including error handling for wrong passwords.

using System;
using Aspose.Cells;

// This example creates a Workbook, writes sample data, applies a password via Workbook.Settings.Password, saves the file as ODS, checks the IsEncrypted flag with FileFormatUtil, and then reloads the file using LoadOptions.Password to demonstrate that the workbook opens only after the correct password is supplied.
class EncryptOdsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add sample data
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Encrypted ODS test");

        // Set the password to encrypt the workbook
        wb.Settings.Password = "mySecret";

        // Save the workbook as ODS
        string filePath = "encrypted_output.ods";
        wb.Save(filePath, SaveFormat.ODS);

        // Verify that the saved file is encrypted
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine("Is file encrypted? " + formatInfo.IsEncrypted);

        // Load the encrypted workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "mySecret";
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // Verify that the data can be read after providing the password
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Loaded cell value: " + cellValue);
    }
}
