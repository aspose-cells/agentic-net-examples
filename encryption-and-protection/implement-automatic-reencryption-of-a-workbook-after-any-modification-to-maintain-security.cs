// Title: Auto‑re‑encrypt an Aspose.Cells workbook after modification in C#
// Description: Load a password‑protected Excel file with Aspose.Cells, edit its content, then automatically reapply the original password (or a stronger algorithm) before saving, ensuring the workbook stays encrypted.
// Keywords: Aspose.Cells | C# workbook encryption | re‑encrypt Excel | load encrypted workbook | set workbook password | strong encryption Aspose | EncryptionType.StrongCryptographicProvider | SetEncryptionOptions | modify protected workbook | auto reencrypt
// Common Searches: how to re‑encrypt an Excel file after editing with Aspose.Cells | Aspose.Cells C# load password protected workbook and save with same password | set strong encryption when saving a workbook using Aspose.Cells | auto reapply workbook password after changes .NET | batch modify encrypted Excel files Aspose.Cells
// Developer Intent: Load a password‑protected workbook, make changes, then automatically apply the original (or stronger) password before saving.
// Use Cases: Edit a protected .xlsx, update cells, and save it encrypted with the same password using Aspose.Cells for .NET. | Upgrade the encryption level (e.g., StrongCryptographicProvider with 128‑bit key) when persisting a modified workbook. | Integrate automatic re‑encryption into a batch process that processes multiple encrypted workbooks.
// AI Prompts: Generate C# code with Aspose.Cells that opens an encrypted Excel file, updates several cells, and saves it using 256‑bit AES encryption with the original password. | Create a method that accepts a file path, password, and a dictionary of cell addresses/values, modifies the workbook, and re‑applies the password with strong encryption. | Explain how Workbook.Settings.Password and SetEncryptionOptions work together to re‑encrypt a workbook after edits in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookAutoReencryptDemoApp
{
    // Load a password‑protected Excel file with Aspose.Cells, edit its content, then automatically reapply the original password (or a stronger algorithm) before saving, ensuring the workbook stays encrypted.
    public class WorkbookAutoReencryptDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the existing encrypted workbook
            string inputPath = "protected.xlsx";

            // Path for the modified workbook
            string outputPath = "protected_modified.xlsx";

            // Original password used to protect the workbook
            string password = "mySecret";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook with the password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Example modification: write a value to a cell
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Modified after re‑encryption");

                // Re‑apply the password to keep the workbook encrypted
                workbook.Settings.Password = password;

                // Optional: specify stronger encryption options
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
