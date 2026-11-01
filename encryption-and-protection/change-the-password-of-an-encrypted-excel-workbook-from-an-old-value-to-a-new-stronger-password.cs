// Title: Change an existing password on an encrypted Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace old password with new password on encrypted Excel file | How to update workbook encryption password using Aspose.Cells LoadOptions | Aspose.Cells .NET change password of password‑protected XLSX workbook | Load encrypted Excel workbook with old password and save with stronger password in C#
// Tags: change workbook password Aspose.Cells | load encrypted Excel workbook C# | set new workbook encryption password .NET | save workbook with updated password Aspose.Cells | Workbook.Settings.Password usage

using System;
using System.IO;
using Aspose.Cells;

// The example loads a password‑protected XLSX file using the original password via LoadOptions, assigns a stronger password to Workbook.Settings.Password, and saves the workbook to a new file with the updated encryption.
class ChangeWorkbookPassword
{
    static void Main()
    {
        try
        {
            // Path to the encrypted workbook
            string inputPath = @"C:\Temp\EncryptedWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Old password used to open the workbook
            string oldPassword = "OldWeakPass123";

            // New stronger password to set
            string newPassword = "NewStrongerPass!@#456";

            // Load the workbook with the old password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the new password for the workbook
            workbook.Settings.Password = newPassword;

            // Save the workbook with the new password
            string outputPath = @"C:\Temp\EncryptedWorkbook_NewPassword.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Password changed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
