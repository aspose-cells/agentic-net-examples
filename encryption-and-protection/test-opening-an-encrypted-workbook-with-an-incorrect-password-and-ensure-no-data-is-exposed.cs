// Title: Test that an encrypted Excel workbook rejects a wrong password with Aspose.Cells for .NET
// Description: Creates a workbook, applies password protection, saves it, then attempts to load the file using an incorrect password. The sample shows the expected exception, confirms that cell values remain inaccessible, and demonstrates password verification via FileFormatUtil.VerifyPassword without opening the workbook.
// Keywords: Aspose.Cells | C# | encrypted workbook | password protection | wrong password | VerifyPassword | exception handling | Excel security | unit test | LoadOptions
// Common Searches: Aspose.Cells open encrypted file with incorrect password | verify Excel password without loading workbook C# | C# test for password‑protected workbook exception | FileFormatUtil VerifyPassword example | prevent data leak when opening protected Excel
// Developer Intent: Confirm that loading a password‑protected workbook with an invalid password fails and does not expose any cell data.
// Use Cases: Automated unit test to validate security of password‑protected Excel files. | Pre‑flight password check before processing an encrypted workbook in a web service. | Logging and graceful handling of unauthorized access attempts to protected spreadsheets.
// AI Prompts: Generate an NUnit test that asserts Aspose.Cells throws an exception for a wrong password and that VerifyPassword returns false. | Provide C# code that safely opens a password‑protected Excel file, handling incorrect passwords without reading sheet data. | Refactor the example into two reusable methods: one for password verification and another for workbook loading with proper error handling.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, applies password protection, saves it, then attempts to load the file using an incorrect password. The sample shows the expected exception, confirms that cell values remain inaccessible, and demonstrates password verification via FileFormatUtil.VerifyPassword without opening the workbook.
class TestEncryptedWorkbook
{
    static void Main()
    {
        // Create a new workbook and add confidential data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Protect the workbook with a known password
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook
        string filePath = "encrypted.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Attempt to open the encrypted workbook with an incorrect password
        try
        {
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "wrongPassword";

            // This line should throw an exception because the password is invalid
            Workbook wbWrong = new Workbook(filePath, loadOptions);

            // If no exception occurs, trying to read data would indicate a security breach
            string leakedValue = wbWrong.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Unexpectedly accessed data: " + leakedValue);
        }
        catch (Exception ex)
        {
            // Expected path: loading fails and no data is exposed
            Console.WriteLine("Failed to open workbook with incorrect password: " + ex.Message);
        }

        // Verify password without loading the workbook's content
        using (Stream stream = File.OpenRead(filePath))
        {
            bool isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, "wrongPassword");
            Console.WriteLine("Password verification (should be false): " + isPasswordCorrect);
        }
    }
}
