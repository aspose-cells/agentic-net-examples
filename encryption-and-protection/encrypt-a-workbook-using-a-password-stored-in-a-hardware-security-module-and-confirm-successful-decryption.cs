// Title: Encrypt an Excel workbook with a password from a hardware security module and validate decryption using Aspose.Cells for .NET
// Common Searches: how to protect an Excel file with a password retrieved from an HSM using Aspose.Cells C# | Aspose.Cells C# load encrypted .xlsx with LoadOptions password | verify that a workbook encrypted with Settings.Password can be decrypted in Aspose.Cells | C# example of using a hardware security module to set workbook password in Aspose.Cells
// Tags: encrypt workbook with HSM password Aspose.Cells | Aspose.Cells workbook password protection C# | LoadOptions password decryption Aspose.Cells | hardware security module password retrieval C# | validate Excel decryption Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample creates a new workbook, obtains a password from a mock HSM method, encrypts the workbook by assigning Settings.Password, saves it as an XLSX file, then reloads the file using LoadOptions with the same password and confirms successful decryption by checking the value of cell A1.
class Program
{
    // Mock method representing retrieval of a password from a hardware security module (HSM)
    static string GetPasswordFromHSM()
    {
        // In a real scenario, this method would interface with the HSM SDK/API
        // to securely retrieve the password. Here we return a placeholder.
        return "SecureHSMPassword123!";
    }

    static void Main()
    {
        // Step 1: Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose!");
        sheet.Name = "DemoSheet";

        // Step 2: Retrieve the encryption password from the HSM
        string hsmPassword = GetPasswordFromHSM();

        // Step 3: Encrypt the workbook using the retrieved password
        workbook.Settings.Password = hsmPassword; // Set workbook password
        string encryptedFilePath = "EncryptedWorkbook.xlsx";
        workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook encrypted and saved to '{encryptedFilePath}'.");

        // Step 4: Load the encrypted workbook using the same password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = hsmPassword // Provide password for decryption
        };
        Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

        // Step 5: Verify successful decryption by reading the cell value
        string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
        if (cellValue == "Hello, Aspose!")
        {
            Console.WriteLine("Decryption successful. Cell A1 contains expected value:");
            Console.WriteLine(cellValue);
        }
        else
        {
            Console.WriteLine("Decryption failed or data corrupted.");
        }
    }
}
