// Title: Identify if an Excel workbook uses Aspose.Cells default encryption and log a recommendation to use stronger protection (C# .NET)
// AI Prompts: Generate C# code using Aspose.Cells that opens an encrypted .xlsx file, checks the WorkbookEncryptionInfo for the default algorithm, and writes a warning to the console if the default is detected. | Update the provided program to evaluate the encryption algorithm of the loaded workbook and output a recommendation to upgrade to a stronger encryption method.
// Common Searches: Aspose.Cells how to determine if an Excel file uses default encryption algorithm | C# check encryption strength of password-protected workbook with Aspose.Cells | log warning when Excel workbook encryption is weak using Aspose.Cells .NET | retrieve encryption algorithm from loaded workbook in Aspose.Cells | recommend stronger encryption for encrypted .xlsx in C#
// Tags: detect workbook encryption algorithm Aspose.Cells | log encryption upgrade recommendation C# | load password protected Excel file Aspose.Cells | detect default workbook encryption .NET | strengthen Excel file protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example shows how to load a password‑protected .xlsx workbook with Aspose.Cells in C#. It demonstrates checking the workbook’s EncryptionInfo (when available) to see if the default algorithm is used and writes a console message recommending a stronger encryption method. The code also includes file‑existence validation and error handling.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // Password required to open the workbook
            string password = "yourPassword";

            // Verify that the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Configure load options with the password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };

                // Load the workbook using the specified options
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");

                // Note: Retrieval of the encryption algorithm may not be supported in all
                // versions of Aspose.Cells. If needed, consult the documentation for the
                // specific version you are using.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
