// Title: Validate that a password‑protected Excel workbook cannot be opened without the password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if opening a password‑protected XLSX throws an exception when no password is supplied | how to test Excel workbook encryption with Aspose.Cells and ensure it requires a password | C# code to verify that a workbook saved with Settings.Password cannot be opened without LoadOptions.Password | catch CellsException for missing password when loading an encrypted Excel file with Aspose.Cells | validate password protection on an Excel file using Aspose.Cells .NET API
// Tags: Aspose.Cells password protection validation | load encrypted XLSX with LoadOptions password | catch CellsException for missing workbook password | verify workbook encryption Aspose.Cells .NET | prevent opening password‑protected Excel without password

using System;
using Aspose.Cells;

// The example creates a workbook, applies a password via Settings.Password, saves it as an encrypted XLSX, then attempts to open the file without providing a password, catching the expected CellsException. It finally opens the workbook with the correct password using LoadOptions to confirm successful decryption.
class WorkbookEncryptionValidation
{
    static void Main()
    {
        // 1. Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // 2. Encrypt the workbook with a password
        wb.Settings.Password = "Secret123";

        // 3. Save the encrypted workbook
        string filePath = "Encrypted.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // 4. Attempt to open the encrypted workbook without providing a password
        try
        {
            // This should throw an exception because the password is required
            Workbook wbWithoutPassword = new Workbook(filePath);
            Console.WriteLine("ERROR: Workbook opened without password (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected path: exception indicates that a password is required
            Console.WriteLine("Validation passed: Cannot open workbook without password.");
            Console.WriteLine($"Exception message: {ex.Message}");
        }

        // 5. Open the workbook with the correct password to confirm it works
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = "Secret123"
        };
        Workbook wbWithPassword = new Workbook(filePath, loadOptions);
        Console.WriteLine("Workbook opened successfully with the correct password.");
    }
}
