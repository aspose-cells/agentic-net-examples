using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class OdsEncryptionDemo
{
    static void Main()
    {
        // Path for the encrypted ODS file
        string encryptedOdsPath = "EncryptedWorkbook.ods";

        // Password to protect the workbook
        string password = "SecurePass123";

        // -------------------- Create and encrypt ODS --------------------
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "DataSheet";
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Set the password for the workbook (this will encrypt the file)
        workbook.Settings.Password = password;

        // Optionally specify encryption algorithm and key length (ignored for ODS but required for XLSX compatibility)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save as ODS using default OdsSaveOptions
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        // Example: set generator type (optional)
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
        workbook.Save(encryptedOdsPath, saveOptions);

        // -------------------- Load and decrypt ODS --------------------
        // Prepare load options with the password
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        loadOptions.Password = password; // Password required to open the encrypted ODS

        // Load the encrypted ODS file
        Workbook loadedWorkbook = new Workbook(encryptedOdsPath, loadOptions);

        // Verify that data is accessible after decryption
        Worksheet loadedSheet = loadedWorkbook.Worksheets["DataSheet"];
        Console.WriteLine("Decrypted data:");
        Console.WriteLine($"A2 = {loadedSheet.Cells["A2"].StringValue}");
        Console.WriteLine($"B2 = {loadedSheet.Cells["B2"].StringValue}");
        Console.WriteLine($"A3 = {loadedSheet.Cells["A3"].StringValue}");
        Console.WriteLine($"B3 = {loadedSheet.Cells["B3"].StringValue}");

        // -------------------- Save decrypted version as XLSX (optional) --------------------
        string decryptedXlsxPath = "DecryptedWorkbook.xlsx";
        loadedWorkbook.Save(decryptedXlsxPath, SaveFormat.Xlsx);

        // Clean up
        workbook.Dispose();
        loadedWorkbook.Dispose();
    }
}