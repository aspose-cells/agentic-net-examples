// Title: Verify that a password‑protected XLSX cannot be opened with Excel 97‑2003 using Aspose.Cells for .NET
// Description: C# sample that creates a workbook, writes a cell, applies a password via Workbook.Settings.Password, saves as XLSX, checks Workbook.Settings.IsEncrypted, then tries to load the file with LoadFormat.Excel97To2003 without a password (expecting an exception), and finally opens it correctly with LoadOptions.Password.
// Keywords: Aspose.Cells | C# | .NET | password protected workbook | XLSX encryption | Excel 97-2003 compatibility | LoadFormat.Excel97To2003 | LoadOptions.Password | Workbook.Settings.IsEncrypted | legacy Excel version | exception handling
// Common Searches: Aspose.Cells verify encrypted XLSX cannot be opened in Excel 97-2003 | C# load password protected XLSX without password throws error | Check Workbook.Settings.IsEncrypted after saving | How to test Excel 97-2003 compatibility for encrypted files | Aspose.Cells load options password example
// Developer Intent: Ensure that a workbook encrypted with a password is rejected when opened as an Excel 97‑2003 file without providing the password.
// Use Cases: Automated build validation to confirm encrypted spreadsheets are inaccessible to legacy Excel versions. | Compliance auditing to guarantee sensitive data remains protected when shared with older Office installations. | Application logic that blocks opening of password‑protected files on unsupported Excel formats.
// AI Prompts: Write C# code using Aspose.Cells that attempts to open a password‑protected XLSX with LoadFormat.Excel97To2003 and captures the expected exception. | Explain the relationship between Workbook.Settings.IsEncrypted, Workbook.Settings.Password, and LoadOptions.Password when handling encrypted files across different Excel versions.

using System;
using Aspose.Cells;

// C# sample that creates a workbook, writes a cell, applies a password via Workbook.Settings.Password, saves as XLSX, checks Workbook.Settings.IsEncrypted, then tries to load the file with LoadFormat.Excel97To2003 without a password (expecting an exception), and finally opens it correctly with LoadOptions.Password.
class VerifyEncryptionOlderVersion
{
    static void Main()
    {
        // Create a new workbook and put some data
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "mySecret";

        // Save the workbook in the modern XLSX format
        string filePath = "encrypted.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Verify that the workbook reports being encrypted
        Console.WriteLine("Workbook.IsEncrypted after save: " + wb.Settings.IsEncrypted);

        // Attempt to open the encrypted file using an older Excel format (Excel 97-2003) without providing a password
        try
        {
            LoadOptions oldFormatOptions = new LoadOptions(LoadFormat.Excel97To2003);
            // No password is supplied here; older versions should not be able to open it
            Workbook oldVersionWb = new Workbook(filePath, oldFormatOptions);
            Console.WriteLine("Unexpectedly opened with older Excel format.");
        }
        catch (Exception ex)
        {
            // Expected path: an exception is thrown because the file is encrypted and older versions cannot handle it
            Console.WriteLine("Failed to open with older Excel version as expected: " + ex.Message);
        }

        // Open the encrypted workbook correctly by providing the password (auto-detect format)
        LoadOptions correctOptions = new LoadOptions();
        correctOptions.Password = "mySecret";
        Workbook openedWb = new Workbook(filePath, correctOptions);
        Console.WriteLine("Opened with password, cell A1 value: " + openedWb.Worksheets[0].Cells["A1"].StringValue);
    }
}
