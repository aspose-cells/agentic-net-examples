// Title: Encrypt an ODS workbook with a password using Aspose.Cells for .NET and confirm the password prompt on open
// AI Prompts: Create a new workbook, assign workbook.Settings.Password, and save it as an encrypted ODS file with Aspose.Cells in C#. | Attempt to load the encrypted ODS file without a password, catch the CellsException, and verify that protection is enforced. | Open a password‑protected ODS workbook by providing the correct password through LoadOptions in Aspose.Cells.
// Common Searches: asp.net encrypt ods workbook password Aspose.Cells example | how to test password protection on ODS file using Aspose.Cells C# | catch CellsException when opening encrypted ODS without password | load password protected ODS with LoadOptions Aspose.Cells .NET
// Tags: encrypt ODS workbook Aspose.Cells | password protection ODS Aspose.Cells | load encrypted ODS with LoadOptions | handle CellsException missing password | set workbook.Settings.Password C#

using System;
using Aspose.Cells;

// The sample creates a workbook, sets a password via workbook.Settings.Password, saves it as an encrypted ODS file, then attempts to open it without a password to catch a CellsException, and finally demonstrates successful opening using LoadOptions with the correct password.
class OdsEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");

        // Set a password to encrypt the ODS file
        // When the file is opened, Aspose.Cells will prompt for this password
        workbook.Settings.Password = "SecretPassword";

        // Save the workbook as ODS with encryption
        string encryptedFile = "EncryptedWorkbook.ods";
        workbook.Save(encryptedFile, SaveFormat.ODS);

        Console.WriteLine($"Workbook saved and encrypted as '{encryptedFile}'.");

        // Verify that opening without a password throws an exception
        try
        {
            // Attempt to load the encrypted file without providing a password
            Workbook loadedWorkbook = new Workbook(encryptedFile);
            // If no exception, the file was not protected (unexpected)
            Console.WriteLine("ERROR: Workbook opened without password. Encryption failed.");
        }
        catch (CellsException ex)
        {
            // Expected exception for password-protected files
            Console.WriteLine("Verification successful: Opening the file without a password throws an exception.");
            Console.WriteLine($"Exception Message: {ex.Message}");
        }

        // Optional: Demonstrate opening with the correct password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.ODS)
        {
            Password = "SecretPassword"
        };
        Workbook openedWithPassword = new Workbook(encryptedFile, loadOptions);
        Console.WriteLine("Workbook opened successfully with the correct password.");
    }
}
