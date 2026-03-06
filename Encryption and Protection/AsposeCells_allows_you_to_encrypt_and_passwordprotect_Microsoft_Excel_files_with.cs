using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook() constructor)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password that will be required to open the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Optionally specify encryption options (for XLS files; ignored for XLSX but shown for completeness)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Protect the workbook structure (prevents adding/removing/renaming sheets) with the same password
            workbook.Protect(ProtectionType.Structure, "StrongPassword123");

            // Save the encrypted and protected workbook as XLSX (uses the provided Save method)
            string outputPath = "EncryptedProtectedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Demonstrate loading the protected workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "StrongPassword123";
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);

            // Verify that the data can be read after providing the correct password
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Loaded cell value: {cellValue}");
        }
    }
}