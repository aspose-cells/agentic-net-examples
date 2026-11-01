// Title: Update the password of an encrypted Excel workbook using Aspose.Cells for .NET and save it as a new file
// Common Searches: asp.net change password of encrypted xlsx workbook using Aspose.Cells | load encrypted Excel file with old password and re‑save with new password in C# | how to retrieve workbook passwords from environment variables in Aspose.Cells .NET | replace workbook protection password programmatically with Aspose.Cells | securely update Excel file encryption password using Aspose.Cells LoadOptions
// Tags: Aspose.Cells change workbook password C# | load encrypted XLSX with password Aspose.Cells | save Excel file with new protection password Aspose.Cells | environment variable password retrieval .NET | secure workbook re‑encryption Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example reads old and new passwords from environment variables, loads an encrypted .xlsx workbook with the old password via Aspose.Cells LoadOptions, assigns a new password to Workbook.Settings.Password, and saves the workbook as a new file, handling missing files and missing environment variables.
class Program
{
    // Environment variable names for passwords (replace with your own method of secure storage)
    private const string OldPasswordEnv = "WORKBOOK_OLD_PASSWORD";
    private const string NewPasswordEnv = "WORKBOOK_NEW_PASSWORD";

    static async Task Main(string[] args)
    {
        // Paths to the workbook files
        string inputFilePath = @"C:\Files\EncryptedWorkbook.xlsx";
        string outputFilePath = @"C:\Files\EncryptedWorkbook_Updated.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputFilePath))
            {
                Console.Error.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            // Retrieve passwords (simulated async for compatibility)
            var (oldPassword, newPassword) = await GetPasswordsAsync();

            // Load the encrypted workbook using the old password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            var workbook = new Workbook(inputFilePath, loadOptions);

            // Change the workbook password to the new one
            workbook.Settings.Password = newPassword;

            // Save the workbook; the new password will be applied automatically
            workbook.Save(outputFilePath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook password has been successfully updated.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retrieves passwords from environment variables; replace with your secure retrieval logic if needed
    private static Task<(string oldPassword, string newPassword)> GetPasswordsAsync()
    {
        string oldPwd = Environment.GetEnvironmentVariable(OldPasswordEnv);
        string newPwd = Environment.GetEnvironmentVariable(NewPasswordEnv);

        if (string.IsNullOrEmpty(oldPwd) || string.IsNullOrEmpty(newPwd))
        {
            throw new InvalidOperationException("Password environment variables are not set.");
        }

        return Task.FromResult((oldPwd, newPwd));
    }
}
