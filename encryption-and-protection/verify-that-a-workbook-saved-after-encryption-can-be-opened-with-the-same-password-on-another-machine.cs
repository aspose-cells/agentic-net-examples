// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and confirm it can be opened on another machine
// AI Prompts: Write C# code that creates a new Workbook, assigns an opening password via wb.Settings.Password, saves it as an .xlsx file, then reloads the file with LoadOptions.Password and reads a cell to verify decryption. | Show how to protect an Excel file with a password in Aspose.Cells and programmatically open the same file on a different computer using the same password.
// Common Searches: how to set opening password for an Excel workbook using Aspose.Cells C# | load password-protected .xlsx with Aspose.Cells LoadOptions in .NET | verify that an encrypted Excel file can be opened on another machine using Aspose.Cells | C# example for saving and reopening a password-protected workbook with Aspose.Cells | Aspose.Cells encrypt workbook and read cell after decryption
// Tags: Aspose.Cells set workbook opening password | C# save encrypted xlsx with Aspose.Cells | LoadOptions password decryption Aspose.Cells | cross-machine workbook password verification | validate encrypted Excel file Aspose.Cells

using System;
using Aspose.Cells;

// // Demonstrates creating a workbook, applying an opening password, saving it, then loading it with LoadOptions.Password to confirm the content can be read after decryption.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        // Define the password to protect the workbook
        const string password = "Secret123";

        // ------------------- Create and encrypt workbook -------------------
        // Create a new workbook
        Workbook wb = new Workbook();

        // Write a test value to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Encrypted Content");

        // Set the opening password (this encrypts the file)
        wb.Settings.Password = password;

        // Save the encrypted workbook to a file
        string filePath = "encrypted.xlsx";
        wb.Save(filePath);

        // ------------------- Load and verify workbook -------------------
        // Load the workbook using the same password
        LoadOptions loadOptions = new LoadOptions
        {
            Password = password
        };
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // Read the value from the loaded workbook to verify successful decryption
        string readValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;

        // Output the verification result
        Console.WriteLine("Successfully opened encrypted workbook.");
        Console.WriteLine("Read value: " + readValue);
    }
}
