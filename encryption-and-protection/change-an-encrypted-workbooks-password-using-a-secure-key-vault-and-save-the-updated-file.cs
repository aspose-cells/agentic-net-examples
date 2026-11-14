// Title: Change the password of an encrypted Excel workbook using Aspose.Cells for .NET and environment variables
// Common Searches: how to replace workbook password in an encrypted Excel file using Aspose.Cells C# | load an encrypted .xlsx with old password and save with new password Aspose.Cells .NET | using environment variables to manage Excel file passwords with Aspose.Cells | update protection of an encrypted workbook programmatically in C# | Aspose.Cells change Excel file password without opening in UI
// Tags: Aspose.Cells change workbook password .NET | load encrypted Excel workbook with password Aspose.Cells | save Excel workbook with new password Aspose.Cells | environment variable password handling C# | programmatic Excel file protection update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads the old and new passwords from environment variables, loads an encrypted XLSX workbook with the original password via LoadOptions, assigns a new password to workbook.Settings.Password, and saves the file as a new encrypted workbook, handling missing variables and file‑not‑found errors.
class Program
{
    static void Main()
    {
        try
        {
            // Retrieve passwords from environment variables (replace with your secure source)
            string oldPassword = Environment.GetEnvironmentVariable("WorkbookOldPassword");
            string newPassword = Environment.GetEnvironmentVariable("WorkbookNewPassword");

            if (string.IsNullOrEmpty(oldPassword) || string.IsNullOrEmpty(newPassword))
            {
                Console.WriteLine("Error: Passwords are not set in environment variables.");
                return;
            }

            // Input workbook path
            string inputPath = "encryptedWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the encrypted workbook using the old password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };

            var workbook = new Workbook(inputPath, loadOptions);

            // Set the new password for the workbook
            workbook.Settings.Password = newPassword;

            // Output workbook path
            string outputPath = "encryptedWorkbook_updated.xlsx";

            // Save the workbook with the updated password
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
