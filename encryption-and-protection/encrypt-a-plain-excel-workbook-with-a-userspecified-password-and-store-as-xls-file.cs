// Title: C# – Password‑Protect and Encrypt an Excel Workbook (XLS) with Aspose.Cells
// Description: Learn how to create a new workbook, apply a user‑defined password, enable 128‑bit strong encryption, save it as an Excel97‑2003 (XLS) file, and verify the protection by loading it with Aspose.Cells for .NET.
// Keywords: Aspose.Cells encrypt workbook C# | password protect XLS Aspose | strong encryption Excel97-2003 | Workbook.Settings.Password | SetEncryptionOptions Aspose | load password protected workbook | C# Excel file security
// Common Searches: how to encrypt an Excel file with a password using Aspose.Cells | save password‑protected XLS with Aspose.Cells .NET | set 128‑bit encryption for Excel workbook C# | load protected XLS file Aspose.Cells example
// Developer Intent: Secure a workbook with a custom password, apply strong encryption, save as XLS, and confirm access via password‑based loading.
// Use Cases: Generate a fresh workbook, insert data, assign Workbook.Settings.Password, call SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128), then save as SaveFormat.Excel97To2003. | Open the encrypted XLS using LoadOptions.Password to read or modify protected cells. | Swap the encryption algorithm (e.g., to AES128) by changing the EncryptionType argument to meet compliance standards.
// AI Prompts: Write C# code that encrypts a new Excel workbook with a user‑provided password, uses 128‑bit strong encryption, and saves it as an XLS file via Aspose.Cells. | Show how to open a password‑protected XLS file with Aspose.Cells and retrieve the value of a specific cell. | Explain how to change the encryption type to AES128 for an Aspose.Cells workbook before saving.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Learn how to create a new workbook, apply a user‑defined password, enable 128‑bit strong encryption, save it as an Excel97‑2003 (XLS) file, and verify the protection by loading it with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // User‑specified password
            string password = "MySecretPassword";

            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted content");

            // Set the password that will protect the workbook (WorkbookSettings.Password)
            workbook.Settings.Password = password;

            // Optional: set stronger encryption options (Workbook.SetEncryptionOptions)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook as an XLS file (lifecycle: save)
            string outputPath = "EncryptedWorkbook.xls";
            workbook.Save(outputPath, SaveFormat.Excel97To2003);

            // Verify by loading the encrypted workbook with the password (lifecycle: load)
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;

            Console.WriteLine($"Loaded cell value: {cellValue}");
        }
    }
}
