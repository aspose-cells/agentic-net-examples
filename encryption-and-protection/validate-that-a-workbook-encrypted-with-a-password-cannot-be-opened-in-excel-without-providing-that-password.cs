// Title: Check that a password‑protected Excel workbook cannot be opened without the password using Aspose.Cells for .NET
// Common Searches: how to confirm Excel file is encrypted with Aspose.Cells and throws error when opened without password in C# | Aspose.Cells C# load encrypted .xlsx without providing password should raise exception | verify password protection of workbook using Aspose.Cells LoadOptions in .NET | test opening password‑protected Excel workbook without password using Aspose.Cells | C# code to ensure encrypted workbook cannot be accessed without correct password Aspose.Cells
// Tags: Aspose.Cells verify password protection | C# load encrypted workbook without password | Aspose.Cells LoadOptions password handling | Excel workbook encryption exception Aspose.Cells | protected workbook access test Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, writes data, applies a password via Workbook.Settings.Password, and saves it. It then attempts to open the file without a password, which throws an exception, confirming the protection. Finally, the workbook is opened successfully using LoadOptions with the correct password, demonstrating proper encryption validation.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "MySecretPassword";

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        workbook.Save(filePath);

        // Attempt to open the workbook without providing a password
        try
        {
            // This should fail because the workbook is password protected
            Workbook withoutPassword = new Workbook(filePath);
            Console.WriteLine("Unexpectedly opened the workbook without a password.");
        }
        catch (Exception ex)
        {
            // Expected path: an exception is thrown
            Console.WriteLine("Failed to open without password as expected: " + ex.Message);
        }

        // Open the workbook with the correct password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MySecretPassword";
        Workbook withPassword = new Workbook(filePath, loadOptions);
        Console.WriteLine("Opened with password successfully. Cell A1 value: " + 
                          withPassword.Worksheets[0].Cells["A1"].StringValue);
    }
}
