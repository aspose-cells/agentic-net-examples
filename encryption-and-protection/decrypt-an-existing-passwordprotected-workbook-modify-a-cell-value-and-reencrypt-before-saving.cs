// Title: Decrypt a password‑protected XLSX workbook, edit a cell, and re‑encrypt with a new password using Aspose.Cells for .NET
// Common Searches: how to open a password protected XLSX file with Aspose.Cells in C# | update a cell in an encrypted Excel workbook and save with a new password using Aspose.Cells | change workbook encryption password programmatically with Aspose.Cells .NET | load protected Excel file, modify data, and re‑save with a different password in C# | Aspose.Cells example for decrypting, editing, and re‑encrypting an Excel file
// Tags: Aspose.Cells load encrypted XLSX | modify cell in encrypted workbook C# | set new workbook password Aspose.Cells | save workbook with re‑encryption .NET | change Excel file encryption Aspose

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the source file, loads a password‑protected XLSX workbook using LoadOptions with the original password, updates cell A1 to "Updated", assigns a new password via workbook.Settings.Password, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Paths to the source and destination workbooks
        string sourcePath = "protected.xlsx";
        string destinationPath = "protected_modified.xlsx";

        // Passwords for opening and re‑encrypting the workbook
        string openPassword = "oldPassword";
        string savePassword = "newPassword";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the password‑protected workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = openPassword // Decrypt using the existing password
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Modify a cell value (example: set A1 to "Updated")
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Updated");

            // Set a new password for the workbook before saving
            workbook.Settings.Password = savePassword; // Encrypt with the new password

            // Save the modified workbook
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
