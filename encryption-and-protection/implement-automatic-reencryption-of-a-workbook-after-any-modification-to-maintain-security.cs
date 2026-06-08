using System;
using Aspose.Cells;

namespace AsposeCellsSecurityDemo
{
    class Program
    {
        // Path to the source workbook (can be encrypted or plain)
        private const string SourcePath = "source.xlsx";

        // Path to the output workbook
        private const string OutputPath = "output_secure.xlsx";

        // Password used for encryption
        private const string WorkbookPassword = "StrongP@ssw0rd";

        static void Main()
        {
            // Load the workbook. If it is already encrypted, provide the password.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = WorkbookPassword; // Works even if the file is not encrypted

            Workbook workbook = new Workbook(SourcePath, loadOptions);

            // ----- Perform any modifications -----
            // Example: write a value to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified at " + DateTime.Now);

            // ----- Re‑encrypt before saving -----
            // Ensure the workbook has the desired password set.
            // If the workbook was previously encrypted, this will re‑apply the same password.
            workbook.Settings.Password = WorkbookPassword;

            // Optionally, set stronger encryption (e.g., AES 128‑bit)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the workbook. The password set above will cause the file to be encrypted.
            workbook.Save(OutputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook saved securely to '{OutputPath}'.");
        }
    }
}