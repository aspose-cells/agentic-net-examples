// Title: Decrypt an Encrypted ODS Workbook, Edit Cells, and Re‑Encrypt with a New Password using Aspose.Cells for .NET
// Description: Load a password‑protected ODS file with Aspose.Cells, change cell A1, assign a new workbook password, and save the file as an ODS document encrypted with the new password.
// Keywords: Aspose.Cells | C# | .NET | ODS | encrypted workbook | password protected spreadsheet | change workbook password | modify cell | LoadOptions.Password | Workbook.Settings.Password | SaveFormat.Ods
// Common Searches: how to open encrypted ODS with Aspose.Cells C# | change password of a protected ODS file using .NET | edit cell in password protected ODS workbook Aspose | re‑save ODS with a new password C# Aspose.Cells | load, modify, and re‑encrypt ODS spreadsheet programmatically
// Developer Intent: Load an encrypted ODS file, update its content, and save it encrypted with a different password.
// Use Cases: Batch‑process confidential ODS reports: decrypt, update status cells, and re‑encrypt with a corporate password. | Automate password rotation for ODS templates after data validation before distribution. | Schedule a task that injects dynamic values into an encrypted ODS file and saves it with a new password for secure sharing.
// AI Prompts: Write C# code with Aspose.Cells that opens an ODS file protected by 'oldPass', sets cell B2 to DateTime.Now, and saves it as ODS encrypted with 'newPass'. | Explain step‑by‑step how Aspose.Cells decrypts an ODS workbook on load and re‑encrypts it on save when different passwords are supplied. | Create a reusable method that takes an input ODS path, old password, new password, and a dictionary of cell addresses with values, then updates and saves the workbook encrypted.

using System;
using Aspose.Cells;

// Load a password‑protected ODS file with Aspose.Cells, change cell A1, assign a new workbook password, and save the file as an ODS document encrypted with the new password.
class Program
{
    static void Main()
    {
        // Input encrypted ODS file and passwords
        string inputFile = "encrypted_input.ods";
        string oldPassword = "oldPass";
        string newPassword = "newPass";

        // Load the workbook with the original password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = oldPassword;
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Modify a cell value (example: set A1 to "Modified")
        workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // Apply a new password for the workbook
        workbook.Settings.Password = newPassword;

        // Save the workbook as ODS with the new password
        workbook.Save("reencrypted_output.ods", SaveFormat.Ods);

        // Clean up
        workbook.Dispose();
    }
}
