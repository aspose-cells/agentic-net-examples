// Title: Encrypt an Excel workbook with a password fetched from a hardware security module and validate decryption using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to assign Workbook.Settings.Password with the password obtained from an HSM and save the workbook as an encrypted .xlsx file. | Open the encrypted workbook by providing LoadOptions.Password set to the HSM‑derived password, then read a known cell to confirm decryption succeeded. | Swap the placeholder GetPasswordFromHSM method for a real HSM SDK call while preserving the same encryption/decryption workflow.
// Common Searches: asp.net retrieve password from hardware security module to encrypt Excel with Aspose.Cells | how to open password protected .xlsx using Aspose.Cells when password is stored in HSM | verify decryption of an encrypted workbook created with Aspose.Cells by reading a cell value | c# example for workbook encryption using dynamic password from external source
// Tags: workbook encryption with Aspose.Cells C# | load password protected Excel using LoadOptions Aspose.Cells | HSM password integration Aspose.Cells .NET | validate decryption by reading cell Aspose.Cells | dynamic password workbook protection Aspose.Cells

using System;
using Aspose.Cells;

// The sample program simulates obtaining a password from a hardware security module, creates a workbook with sample data, encrypts it by setting Workbook.Settings.Password, saves it as EncryptedWorkbook.xlsx, then reloads the file using LoadOptions.Password and reads cell B2 to confirm that decryption was successful.
class Program
{
    // Simulated method that retrieves the password from a hardware security module (HSM)
    static string GetPasswordFromHSM()
    {
        // In a real scenario, this method would interface with the HSM SDK/API
        // to securely retrieve the password. Here we return a placeholder.
        return "SecureHSMPassword123!";
    }

    static void Main()
    {
        // Retrieve the encryption password from the HSM
        string password = GetPasswordFromHSM();

        // ------------------- Create Workbook -------------------
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "DataSheet";
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // ------------------- Encrypt Workbook -------------------
        // Set the password for encryption
        workbook.Settings.Password = password;

        // Save the encrypted workbook to a file
        string encryptedFilePath = "EncryptedWorkbook.xlsx";
        workbook.Save(encryptedFilePath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook encrypted and saved to '{encryptedFilePath}'.");

        // ------------------- Load and Decrypt Workbook -------------------
        // Load the workbook using the same password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };
        Workbook loadedWorkbook = new Workbook(encryptedFilePath, loadOptions);

        // Verify decryption by reading a known cell value
        Worksheet loadedSheet = loadedWorkbook.Worksheets["DataSheet"];
        string nameInCellB2 = loadedSheet.Cells["B2"].StringValue;

        Console.WriteLine($"Decryption successful. Value in B2: {nameInCellB2}");
    }
}
