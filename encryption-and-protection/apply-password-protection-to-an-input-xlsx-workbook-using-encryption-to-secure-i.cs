using System;
using Aspose.Cells;

namespace AsposeCellsPasswordProtection
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that needs to be protected
            string inputPath = "input.xlsx";

            // Path where the password‑protected workbook will be saved
            string outputPath = "output_protected.xlsx";

            // Password to encrypt the workbook
            string password = "mySecret";

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply encryption password (this will require the password to open the file)
            workbook.Settings.Password = password;

            // Save the workbook with encryption
            workbook.Save(outputPath);

            // Optional: Verify that the workbook is encrypted by loading it with the password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook protectedWorkbook = new Workbook(outputPath, loadOptions);

            // Output a simple verification message
            Console.WriteLine("Workbook encrypted: " + protectedWorkbook.Settings.IsEncrypted);
        }
    }
}