// Title: Open a password‑protected Excel workbook using Aspose.Cells LoadOptions in C#
// AI Prompts: Load an encrypted .xlsx file by setting LoadOptions.Password, then read a specific cell value. | Decrypt an Excel workbook on opening with Aspose.Cells and optionally save it as an unprotected file.
// Common Searches: C# Aspose.Cells load encrypted Excel file with password | How to read a cell from a password protected workbook using Aspose.Cells .NET | Aspose.Cells LoadOptions Password property example for .xlsx | Remove protection from Excel file while loading with Aspose.Cells C# | Error handling for wrong password when opening Excel with Aspose.Cells
// Tags: load encrypted xlsx with Aspose.Cells LoadOptions | Aspose.Cells password protected workbook opening | read cell from protected Excel C# | save decrypted workbook Aspose.Cells | exception handling invalid password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    // The example checks for the encrypted Excel file, creates a LoadOptions object with the required password, loads the workbook using Aspose.Cells, reads the value of cell A1 from the first worksheet, and optionally saves a decrypted copy. It also demonstrates proper exception handling for missing files or incorrect passwords.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string filePath = "EncryptedWorkbook.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Set load options with the password for the encrypted workbook
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = "YourPasswordHere"
                };

                // Load the workbook using the specified load options
                var workbook = new Workbook(filePath, loadOptions);

                // Read the value of cell A1 from the first worksheet
                var sheet = workbook.Worksheets[0];
                string cellValue = sheet.Cells["A1"].StringValue;
                Console.WriteLine($"Value in A1: {cellValue}");

                // (Optional) Save the workbook to a new file without encryption
                // workbook.Save("DecryptedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                // Handle any errors, including invalid password or file issues
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
