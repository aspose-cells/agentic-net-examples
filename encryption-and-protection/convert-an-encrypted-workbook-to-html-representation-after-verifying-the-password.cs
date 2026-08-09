// Title: Convert Password‑Protected Excel to HTML with Aspose.Cells for .NET
// Description: A C# example that checks the existence of an encrypted .xlsx file, verifies the supplied password using FileFormatUtil.VerifyPassword, loads the workbook with LoadOptions, and saves it as an HTML document via SaveFormat.Html, with graceful error handling for missing files or wrong passwords.
// Keywords: Aspose.Cells | .NET | encrypted workbook | verify password | password protected Excel | convert to HTML | LoadOptions | FileFormatUtil | C# example | save as HTML
// Common Searches: Aspose.Cells open password protected Excel file | verify Excel workbook password before loading Aspose.Cells | convert encrypted .xlsx to HTML C# | Aspose.Cells load options password example | how to export protected Excel to HTML using Aspose
// Developer Intent: Validate the Excel file password and export the protected workbook to HTML.
// Use Cases: Check a user‑provided password before generating an HTML preview of a secured spreadsheet. | Batch‑process multiple password‑protected Excel files, confirming each password and converting them to HTML for web publishing. | Display a clear error message when the supplied password is incorrect while attempting the conversion.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify a password for an encrypted .xlsx file and, if valid, saves the workbook as HTML. | Explain best practices for handling incorrect passwords when converting a protected Excel workbook to HTML with Aspose.Cells. | Show how to iterate over a folder of encrypted Excel files, verify each password, and convert each file to HTML using Aspose.Cells in C#.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# example that checks the existence of an encrypted .xlsx file, verifies the supplied password using FileFormatUtil.VerifyPassword, loads the workbook with LoadOptions, and saves it as an HTML document via SaveFormat.Html, with graceful error handling for missing files or wrong passwords.
    class EncryptedWorkbookToHtml
    {
        static void Main()
        {
            try
            {
                // Path to the encrypted workbook and the password to open it
                string encryptedFilePath = "encrypted.xlsx";
                string password = "myPassword";

                // Verify that the file exists
                if (!File.Exists(encryptedFilePath))
                {
                    Console.WriteLine($"File not found: {encryptedFilePath}");
                    return;
                }

                // Verify the password before loading the workbook
                bool passwordIsCorrect;
                using (FileStream stream = File.OpenRead(encryptedFilePath))
                {
                    passwordIsCorrect = FileFormatUtil.VerifyPassword(stream, password);
                }

                if (!passwordIsCorrect)
                {
                    Console.WriteLine("The provided password is incorrect.");
                    return;
                }

                // Load the encrypted workbook using LoadOptions with the verified password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

                // Convert the workbook to HTML format
                string htmlOutputPath = "output.html";
                workbook.Save(htmlOutputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook successfully saved as HTML: {htmlOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
