// Title: Replace the opening password of an encrypted Excel workbook using Aspose.Cells for .NET without decrypting the file
// AI Prompts: Generate C# code that opens a password‑protected .xlsx with Aspose.Cells, assigns a new opening password, and saves the workbook while preserving its encryption. | Show how to use LoadOptions and Workbook.Settings.Password in Aspose.Cells to update an Excel file’s password without recreating the workbook.
// Common Searches: asp.net change password of encrypted xlsx using aspose.cells loadoptions | c# replace opening password of password protected excel file without decrypting | how to update workbook password in Aspose.Cells without losing data | programmatically modify excel file password with Aspose.Cells .NET
// Tags: Aspose.Cells change workbook opening password | LoadOptions password protected Excel .NET | Workbook.Settings.Password update | Encrypt Excel file with Aspose.Cells | C# replace .xlsx password without decryption

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordExample
{
    // // Demonstrates creating an Excel workbook, encrypting it with an opening password, loading it using LoadOptions, changing the password via Workbook.Settings.Password, and saving the file again—all with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // 2. Set an opening password (encrypt the workbook)
                workbook.Settings.Password = "oldPass";

                // 3. Save the encrypted workbook to disk
                string filePath = "encrypted.xlsx";
                workbook.Save(filePath, SaveFormat.Xlsx);

                // 4. Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // 5. Load the encrypted workbook using the original password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = "oldPass"
                };
                Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

                // 6. Change the password by assigning a new one
                loadedWorkbook.Settings.Password = "newPass";

                // 7. Save the workbook back, overwriting the original file
                loadedWorkbook.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine("Password changed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
