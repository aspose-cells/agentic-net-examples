// Title: Validate that an Aspose.Cells workbook encrypted with a password cannot be opened without the password and can be opened when the correct password is supplied (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, assign a password, save it, then attempt to open it without supplying the password and handle the resulting error. | Demonstrate how to open a password‑protected Excel file in C# by passing the password through LoadOptions after confirming that an unauthenticated load fails.
// Common Searches: Aspose.Cells C# test opening encrypted Excel file without password | how to catch exception when loading password protected workbook with Aspose.Cells | verify workbook password protection using Aspose.Cells LoadOptions C# | C# code sample for encrypting Excel file and validating password requirement with Aspose.Cells
// Tags: Aspose.Cells encrypt workbook with password | C# load password protected Excel using LoadOptions | exception handling for encrypted workbook Aspose.Cells | validate workbook protection Aspose.Cells | test password requirement Excel .NET

using Aspose.Cells;
using System;

// // Creates a workbook, applies a password for encryption, saves it, then tries to open it without providing the password (expected failure), catches the exception, and finally opens the same file successfully by supplying the correct password via LoadOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Confidential Information");
        sheet.Cells["A2"].PutValue(98765);

        // Encrypt the workbook with a password
        workbook.Settings.Password = "StrongPassword123";

        // Save the encrypted workbook
        string encryptedPath = "EncryptedWorkbook.xlsx";
        workbook.Save(encryptedPath);

        // Attempt to open the encrypted workbook without providing the password
        try
        {
            // This should fail because the password is not supplied
            Workbook withoutPassword = new Workbook(encryptedPath);
            Console.WriteLine("Unexpectedly opened the encrypted workbook without a password.");
        }
        catch (Exception ex)
        {
            // Expected failure – the viewer cannot decrypt the file
            Console.WriteLine("Failed to open encrypted workbook without password: " + ex.Message);
        }

        // Demonstrate successful opening when the correct password is supplied
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "StrongPassword123";
        Workbook withPassword = new Workbook(encryptedPath, loadOptions);
        Console.WriteLine("Workbook opened successfully with the correct password.");
    }
}
