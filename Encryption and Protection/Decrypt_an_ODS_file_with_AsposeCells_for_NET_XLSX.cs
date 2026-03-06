using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDecryptOds
{
    class Program
    {
        static void Main()
        {
            // Path to the encrypted ODS file
            string encryptedOdsPath = Path.Combine(Directory.GetCurrentDirectory(), "encrypted_file.ods");

            if (!File.Exists(encryptedOdsPath))
            {
                Console.WriteLine($"File not found: {encryptedOdsPath}");
                return;
            }

            // Password used to protect the ODS file
            string password = "myPassword";

            // Create load options for ODS and set the password
            OdsLoadOptions loadOptions = new OdsLoadOptions
            {
                Password = password
            };

            // Load the encrypted ODS workbook using the load options
            Workbook workbook = new Workbook(encryptedOdsPath, loadOptions);

            // Example: read a cell value to verify successful decryption
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Decrypted cell A1 value: " + cellValue);

            // Save the workbook as an unprotected XLSX file
            string outputXlsxPath = Path.Combine(Directory.GetCurrentDirectory(), "decrypted_output.xlsx");
            workbook.Save(outputXlsxPath, SaveFormat.Xlsx);

            Console.WriteLine("Decrypted workbook saved to: " + outputXlsxPath);
        }
    }
}