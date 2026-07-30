// Title: C# – Convert a Password‑Protected Excel Workbook to HTML with Aspose.Cells
// Description: Shows how to validate an encrypted workbook’s password using FileFormatUtil.VerifyPassword, load the file with LoadOptions.Password, and export it to an HTML page via HtmlSaveOptions, including robust error handling.
// Keywords: Aspose.Cells C# | encrypted Excel workbook | verify password Aspose.Cells | FileFormatUtil VerifyPassword | LoadOptions Password | HtmlSaveOptions | Excel to HTML conversion | password‑protected .xlsx | convert .xlsx to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells verify password before opening workbook | C# convert password protected Excel to HTML | How to open encrypted .xlsx with Aspose.Cells | Export protected Excel file to HTML using Aspose.Cells | FileFormatUtil VerifyPassword example C#
// Developer Intent: Open a password‑protected Excel file, confirm the supplied password, and generate an HTML representation of the workbook.
// Use Cases: Authenticate a user‑provided password, then render a confidential spreadsheet as a web‑ready HTML preview. | Automate batch conversion of encrypted reports to HTML after successful decryption for publishing on intranet portals.
// AI Prompts: Generate C# code that uses Aspose.Cells to check a workbook password with FileFormatUtil, load the file, and save it as HTML. | Explain how to handle an invalid password when converting a protected Excel file to HTML using Aspose.Cells. | Provide a step‑by‑step guide for converting an encrypted .xlsx to HTML in a .NET console application.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to validate an encrypted workbook’s password using FileFormatUtil.VerifyPassword, load the file with LoadOptions.Password, and export it to an HTML page via HtmlSaveOptions, including robust error handling.
class EncryptedWorkbookToHtml
{
    static void Main()
    {
        // Path to the encrypted workbook and the password to open it
        string encryptedFilePath = "encrypted.xlsx";
        string password = "myPassword";

        // Verify that the input file exists
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"File not found: {encryptedFilePath}");
            return;
        }

        try
        {
            // Verify that the password is correct for the encrypted file (optional)
            bool isPasswordValid;
            using (FileStream stream = File.OpenRead(encryptedFilePath))
            {
                isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
            }

            if (!isPasswordValid)
            {
                Console.WriteLine("Invalid password. Cannot open the workbook.");
                return;
            }

            // Load the encrypted workbook using LoadOptions with the correct password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Convert the workbook to HTML format
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            string htmlOutputPath = "output.html";
            workbook.Save(htmlOutputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {htmlOutputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
