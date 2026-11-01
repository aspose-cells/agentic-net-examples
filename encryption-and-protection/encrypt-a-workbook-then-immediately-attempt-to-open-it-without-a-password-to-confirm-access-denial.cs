// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and verify that opening it without the password fails
// Common Searches: how to encrypt an Excel file with a password using Aspose.Cells C# | Aspose.Cells load encrypted workbook without password exception | C# verify password protection on saved .xlsx using Aspose.Cells | using LoadOptions to open password‑protected Excel in Aspose.Cells .NET | catch access denied error when opening encrypted workbook Aspose.Cells
// Tags: Workbook.Settings.Password property Aspose.Cells | LoadOptions with password for encrypted .xlsx Aspose.Cells | handling access denied exception Aspose.Cells | validate Excel file encryption C# Aspose.Cells | password‑protected workbook loading error Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, writes sample data, sets a password via Workbook.Settings.Password, saves it as an encrypted .xlsx file, then attempts to load the file without providing a password, catches the expected access‑denied exception, and finally demonstrates successful loading using LoadOptions with the correct password.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Encrypt the workbook with a password
        workbook.Settings.Password = "Secret123";
        string encryptedPath = "EncryptedWorkbook.xlsx";
        workbook.Save(encryptedPath, SaveFormat.Xlsx);

        // Attempt to open the encrypted workbook without providing a password
        try
        {
            // This should throw an exception because the password is required
            Workbook openedWithoutPassword = new Workbook(encryptedPath);
            Console.WriteLine("Workbook opened without password (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected outcome: access denied due to missing password
            Console.WriteLine("Access denied as expected: " + ex.Message);
        }

        // Optional: demonstrate successful opening with the correct password
        try
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "Secret123"
            };
            Workbook openedWithPassword = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine("Workbook opened successfully with password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open with password: " + ex.Message);
        }
    }
}
