// Title: C# – Verify Password of Encrypted Excel Workbook and Convert to HTML with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells to confirm a password on an encrypted .xlsx file, load the workbook with LoadOptions, and export it as an HTML page.
// Keywords: Aspose.Cells | C# | encrypted Excel workbook | verify password | FileFormatUtil.VerifyPassword | LoadOptions.Password | save as HTML | Excel to HTML conversion | password‑protected spreadsheet | document security
// Common Searches: Aspose.Cells verify encrypted workbook password C# | convert password protected xlsx to html | load encrypted Excel file with Aspose.Cells | check Excel file password before opening | export protected workbook to html using .NET
// Developer Intent: Confirm the workbook password, then load and export the protected Excel file to HTML.
// Use Cases: Validate user‑supplied passwords in an automated data‑processing pipeline before converting confidential reports to web‑ready HTML. | Provide a secure preview of password‑protected Excel dashboards by converting them to HTML after successful authentication. | Integrate password verification and HTML export into a document‑management system that stores and displays protected spreadsheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify an encrypted Excel file password and then saves the workbook as HTML. | Explain how to handle an invalid password when opening a protected workbook with Aspose.Cells in .NET. | Show how to configure LoadOptions.Password and customize HTML export settings in Aspose.Cells.

using System;
using Aspose.Cells;
using System.IO;

// Demonstrates how to use Aspose.Cells to confirm a password on an encrypted .xlsx file, load the workbook with LoadOptions, and export it as an HTML page.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string encryptedFilePath = "encrypted.xlsx";

        // Password to open the workbook
        string password = "myPassword";

        // Verify that the password is correct before attempting to load
        bool isPasswordValid;
        using (Stream stream = File.OpenRead(encryptedFilePath))
        {
            isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
        }

        if (!isPasswordValid)
        {
            Console.WriteLine("Invalid password. Cannot open the workbook.");
            return;
        }

        // Load the encrypted workbook using the verified password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

        // Convert the workbook to HTML format
        string htmlOutputPath = "output.html";
        workbook.Save(htmlOutputPath, SaveFormat.Html);

        Console.WriteLine($"Workbook successfully converted to HTML: {htmlOutputPath}");
    }
}
