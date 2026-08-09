// Title: Aspose.Cells .NET – Verify encrypted workbook cannot be opened with a wrong password
// Description: Creates an Excel workbook, protects it with a password, saves it, then attempts to load it using an incorrect password via LoadOptions. The example catches the expected exception, confirms that no cell data is exposed, and demonstrates password validation with FileFormatUtil.VerifyPassword without fully loading the file.
// Keywords: Aspose.Cells | C# | .NET | encrypted workbook | password protection | LoadOptions wrong password | FileFormatUtil.VerifyPassword | exception handling | data security | Excel file encryption | US developers | European .NET community
// Common Searches: how to test wrong password on an encrypted Excel file using Aspose.Cells | verify Excel workbook password without opening the file in C# | Aspose.Cells catch exception when loading protected workbook with bad password | prevent data leakage from encrypted Excel when password is incorrect | C# code to check Excel file password validity with Aspose
// Developer Intent: Ensure that an encrypted Excel workbook remains inaccessible and does not reveal any data when an incorrect password is supplied.
// Use Cases: Catch the exception thrown by new Workbook(filePath, loadOptions) to block unauthorized access. | Use FileFormatUtil.VerifyPassword to quickly test password validity before loading large workbooks. | Automate security regression tests that confirm encrypted files protect confidential data.
// AI Prompts: Generate a C# unit test with Aspose.Cells that asserts loading an encrypted workbook with a wrong password throws an exception and returns no cell values. | Show sample code that uses FileFormatUtil.VerifyPassword to programmatically reject an invalid password before opening the workbook. | Create a logging snippet that records a custom error when a workbook fails to open due to an incorrect password while guaranteeing no data is exposed.

using System;
using System.IO;
using Aspose.Cells;

// Creates an Excel workbook, protects it with a password, saves it, then attempts to load it using an incorrect password via LoadOptions. The example catches the expected exception, confirms that no cell data is exposed, and demonstrates password validation with FileFormatUtil.VerifyPassword without fully loading the file.
class TestEncryptedWorkbook
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Protect the workbook with a password
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook
        string filePath = "encrypted.xlsx";
        wb.Save(filePath);

        // Attempt to open the workbook with an incorrect password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "wrongPassword";

        try
        {
            Workbook wbWrong = new Workbook(filePath, loadOptions);
            // If loading succeeds, check if data is exposed (it should not)
            string cellValue = wbWrong.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Unexpectedly opened workbook. Cell value: " + cellValue);
        }
        catch (Exception ex)
        {
            // Expected outcome: loading fails due to wrong password
            Console.WriteLine("Failed to open workbook with incorrect password: " + ex.Message);
        }

        // Verify password without loading the entire workbook
        using (Stream stream = File.OpenRead(filePath))
        {
            bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, "wrongPassword");
            Console.WriteLine("Password verification (should be false): " + isPasswordValid);
        }
    }
}
