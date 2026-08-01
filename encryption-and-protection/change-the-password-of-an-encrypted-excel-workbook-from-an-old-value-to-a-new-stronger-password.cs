// Title: Change the password of an encrypted Excel workbook using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open a password‑protected Excel file with LoadOptions.Password, assign a stronger password via Workbook.Settings.Password, and save the workbook to re‑encrypt it, including file‑existence validation and Aspose.Cells exception handling.
// Keywords: Aspose.Cells | C# | change Excel password | update workbook encryption | LoadOptions.Password | Workbook.Settings.Password | encrypted Excel file | password replacement .NET | Excel file protection | programmatic password change
// Common Searches: Aspose.Cells change workbook password C# | replace Excel file password programmatically | load encrypted workbook with old password .NET | save Excel workbook with new password Aspose | update Excel encryption using C#
// Developer Intent: Replace an existing workbook password with a new, stronger one programmatically.
// Use Cases: Open a password‑protected workbook using the current password. | Set Workbook.Settings.Password to a new value. | Save the workbook so it is re‑encrypted with the new password. | Validate that the source file exists before attempting to load it. | Catch CellsException and generic exceptions to handle errors gracefully.
// AI Prompts: Generate C# code that uses Aspose.Cells to change the password of an encrypted Excel workbook, including error handling. | Provide a reusable method that accepts inputPath, oldPassword, newPassword and updates the workbook’s encryption. | Explain the role of LoadOptions.Password versus Workbook.Settings.Password when re‑encrypting an Excel file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace ChangeWorkbookPasswordDemo
{
    // Demonstrates how to open a password‑protected Excel file with LoadOptions.Password, assign a stronger password via Workbook.Settings.Password, and save the workbook to re‑encrypt it, including file‑existence validation and Aspose.Cells exception handling.
    class Program
    {
        static void Main()
        {
            // Path to the existing encrypted workbook
            string inputPath = "EncryptedWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Old password used to protect the workbook
            string oldPassword = "oldPass123";

            // New, stronger password to replace the old one
            string newPassword = "NewStrongPass!@#456";

            try
            {
                // Load the workbook using the old password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword // password for opening the encrypted file
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Change the encryption password to the new value
                workbook.Settings.Password = newPassword; // password for saving the file

                // Save the workbook; it will be encrypted with the new password
                string outputPath = "EncryptedWorkbook_NewPassword.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine("Password changed successfully.");
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
