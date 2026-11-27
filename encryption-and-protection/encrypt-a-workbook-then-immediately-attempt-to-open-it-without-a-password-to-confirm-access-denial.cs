// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and verify that opening it without a password throws a CellsException
// Common Searches: how to set a password on an Excel workbook with Aspose.Cells in C# | load a password‑protected .xlsx file without providing a password using Aspose.Cells | catch CellsException when opening an encrypted workbook in .NET | verify that a workbook saved with wb.Settings.Password cannot be opened without the password | Aspose.Cells example for encrypting and testing access denial of an Excel file
// Tags: set workbook password Aspose.Cells | save encrypted Excel file Aspose.Cells | open encrypted workbook without password Aspose.Cells | handle CellsException on workbook load | validate workbook encryption Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, writes data to cell A1, assigns a password via wb.Settings.Password, saves the file as an encrypted .xlsx, then attempts to load the same file without supplying a password. The resulting CellsException is caught and its message is displayed, confirming that access is denied without the correct password.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");

        // Encrypt the workbook with a password
        wb.Settings.Password = "Secret123";
        string encryptedPath = "EncryptedWorkbook.xlsx";
        wb.Save(encryptedPath);

        // Attempt to open the encrypted workbook without providing a password
        try
        {
            // This should fail because no password is supplied
            Workbook wbWithoutPassword = new Workbook(encryptedPath);
            Console.WriteLine("Workbook opened without password (unexpected).");
        }
        catch (CellsException ex)
        {
            // Expected outcome: access denied due to missing/incorrect password
            Console.WriteLine("Access denied as expected: " + ex.Message);
        }
    }
}
