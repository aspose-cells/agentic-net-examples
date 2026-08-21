// Title: Validate Excel Workbook Encryption with Aspose.Cells for .NET
// Description: Creates a workbook, applies a password, saves it as XLSX, uses FileFormatUtil.DetectFileFormat to confirm the file is encrypted, then reloads it with LoadOptions.Password and checks Workbook.Settings.IsEncrypted.
// Keywords: Aspose.Cells | C# | Excel encryption | FileFormatUtil | IsEncrypted | LoadOptions password | detect encrypted workbook | validate workbook protection | programmatic encryption check
// Common Searches: Aspose.Cells check if Excel file is password protected | detect encrypted XLSX using Aspose.Cells .NET | verify workbook encryption after saving with Aspose | load password protected Excel with Aspose.Cells | FileFormatUtil IsEncrypted example
// Developer Intent: Confirm that a workbook saved with a password is encrypted and can be opened using the same password.
// Use Cases: Programmatically determine whether a saved Excel file is encrypted via FileFormatUtil.DetectFileFormat. | Open a password‑protected workbook with LoadOptions.Password and verify Settings.IsEncrypted returns true. | Automated testing: assert that encryption flag is true before and after loading the workbook.
// AI Prompts: Generate C# code that opens an existing XLSX file, uses FileFormatUtil to detect encryption, and prints the result. | Write a unit test in C# that saves a workbook with a password, checks FileFormatInfo.IsEncrypted, then loads it with LoadOptions and asserts Settings.IsEncrypted is true. | Provide a step‑by‑step tutorial for validating encryption of an Excel file created with Aspose.Cells, covering detection and loading with the correct password.

using System;
using Aspose.Cells;

// Creates a workbook, applies a password, saves it as XLSX, uses FileFormatUtil.DetectFileFormat to confirm the file is encrypted, then reloads it with LoadOptions.Password and checks Workbook.Settings.IsEncrypted.
class ValidateEncryption
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted workbook test");

        // Set a password to encrypt the workbook
        workbook.Settings.Password = "mySecretPassword";

        // Save the encrypted workbook to a file
        string filePath = "encrypted.xlsx";
        workbook.Save(filePath);

        // Detect encryption status using FileFormatUtil
        FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine("Is the file encrypted? (FileFormatInfo) " + fileInfo.IsEncrypted);

        // Load the workbook with the correct password to verify it can be opened
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.Password = "mySecretPassword";
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
        Console.WriteLine("Workbook loaded successfully. Settings.IsEncrypted: " + loadedWorkbook.Settings.IsEncrypted);
    }
}
