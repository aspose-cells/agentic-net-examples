// Title: Encrypt an Excel workbook with 256‑bit strength using Aspose.Cells for .NET and verify protection
// Description: Demonstrates how to create a workbook, assign a password, apply a 256‑bit strong cryptographic provider via SetEncryptionOptions, save the file, reload it with LoadOptions, and confirm encryption using the Settings.IsEncrypted flag.
// Keywords: Aspose.Cells encryption .NET | 256‑bit Excel workbook protection | SetEncryptionOptions strong provider | Workbook password validation | IsEncrypted flag Aspose | LoadOptions password Excel | C# Excel file security | global data protection | US developers Aspose.Cells
// Common Searches: how to set 256‑bit encryption in Aspose.Cells | verify Excel workbook is encrypted with Aspose | Aspose.Cells SetEncryptionOptions example | C# encrypt Excel file password protection | check if workbook is encrypted Aspose.Cells
// Developer Intent: Apply a 256‑bit password‑based encryption to an Excel workbook with Aspose.Cells and programmatically confirm that the protection is active.
// Use Cases: Secure confidential spreadsheets before distribution. | Automate compliance‑driven encryption for generated reports. | Validate encryption status in CI pipelines to ensure data safety.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, sets a password, uses SetEncryptionOptions with a 256‑bit key, saves the file, reloads it, and checks Settings.IsEncrypted. | Write a reusable method that encrypts an existing .xlsx file with a strong cryptographic provider (256‑bit) using Aspose.Cells and returns true if the file is confirmed encrypted. | Provide a step‑by‑step tutorial for confirming workbook encryption in Aspose.Cells by loading the file with the correct password and inspecting the IsEncrypted property.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Demonstrates how to create a workbook, assign a password, apply a 256‑bit strong cryptographic provider via SetEncryptionOptions, save the file, reload it with LoadOptions, and confirm encryption using the Settings.IsEncrypted flag.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Strength Test");

            // ---------- Set password and encryption strength ----------
            // Password required to open the workbook
            workbook.Settings.Password = "SecretPwd";

            // Set encryption options: use strong provider with 256‑bit key
            // For Excel 2007+ the EncryptionType is ignored, but the key length is applied.
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // ---------- Save the encrypted workbook ----------
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);               // save

            // ---------- Load the workbook with the password ----------
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "SecretPwd"
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions); // load

            // ---------- Validate that the workbook is encrypted ----------
            bool isEncrypted = loadedWorkbook.Settings.IsEncrypted;
            Console.WriteLine($"Workbook is encrypted: {isEncrypted}");

            // Since Aspose.Cells does not expose the actual key length,
            // we confirm encryption by checking IsEncrypted flag.
            // If needed, additional validation can be performed by attempting
            // to open the file with an incorrect password (which would throw).

            // Clean up
            workbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}
