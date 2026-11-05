// Title: Check that a password‑protected Excel workbook cannot be opened without the password using Aspose.Cells for .NET
// Common Searches: how to confirm encrypted Excel file throws exception when opened without password in Aspose.Cells C# | Aspose.Cells C# load encrypted .xlsx without providing a password | verify that workbook password protection blocks unauthorized opening in .NET | exception handling for opening password‑protected workbook with Aspose.Cells | unit test for workbook encryption using Aspose.Cells LoadOptions
// Tags: Aspose.Cells password protection validation | C# load encrypted workbook with LoadOptions | encrypted Excel workbook exception handling | verify workbook encryption Aspose.Cells | password protected .xlsx opening test

using System;
using Aspose.Cells;

// The example creates a workbook, encrypts it with a password, saves it, then attempts to open it without the password to ensure an exception is thrown, and finally opens it successfully by supplying the correct password via LoadOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Secret Data");
        ws.Cells["A2"].PutValue(12345);

        // Encrypt the workbook with a password
        wb.Settings.Password = "MySecretPassword";

        // Save the encrypted workbook
        string encryptedPath = "EncryptedWorkbook.xlsx";
        wb.Save(encryptedPath, SaveFormat.Xlsx);

        // Simulate a third‑party viewer that does not support decryption
        // Attempt to open the file without providing a password
        try
        {
            Workbook wbWithoutPassword = new Workbook(encryptedPath);
            Console.WriteLine("Unexpectedly opened encrypted workbook without password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open encrypted workbook without password: " + ex.Message);
        }

        // Verify that the workbook can be opened when the correct password is supplied
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MySecretPassword";
        Workbook wbWithPassword = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine("Workbook opened successfully with correct password.");
    }
}
