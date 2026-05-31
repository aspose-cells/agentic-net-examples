using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Paths and passwords
            string inputFile = "encrypted_input.ods";
            string oldPassword = "oldPass123";
            string newPassword = "newPass456";
            string outputFile = "reencrypted_output.ods";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the encrypted workbook using the old password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Example modification: change A1 in the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Updated Value");

                // Apply a new password for encryption
                workbook.Settings.Password = newPassword;

                // Save the workbook as ODS with the new encryption
                workbook.Save(outputFile, SaveFormat.Ods);

                Console.WriteLine($"Workbook re‑encrypted and saved to {outputFile}");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}