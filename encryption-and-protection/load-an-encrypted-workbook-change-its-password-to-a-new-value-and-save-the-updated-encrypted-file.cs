// Title: Replace the password of an encrypted Excel workbook using Aspose.Cells for .NET
// AI Prompts: Open a password‑protected .xlsx file with its current password, assign a new password, and save the workbook while preserving encryption using Aspose.Cells in C#. | Create LoadOptions with the existing workbook password, load the workbook, set Workbook.Settings.Password to the desired new value, and write the file to a new encrypted Excel document.
// Common Searches: c# Aspose.Cells change password of a protected Excel file | load encrypted xlsx with old password and save with new password using Aspose.Cells | how to re‑encrypt an Excel workbook with a different password in .NET | Aspose.Cells LoadOptions password parameter example for updating workbook encryption
// Tags: Aspose.Cells load encrypted workbook | Workbook.Settings.Password set new encryption | re‑encrypt Excel file with Aspose.Cells | LoadOptions with workbook password | save workbook using updated password

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an encrypted Excel file, opens it with the current password via LoadOptions, assigns a new password through Workbook.Settings.Password, and saves the workbook to a new file, handling both Aspose.Cells‑specific and general exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook to be processed
        string inputFile = "encrypted_input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        // Current password of the workbook
        string oldPassword = "oldPassword";

        // Desired new password
        string newPassword = "newPassword";

        try
        {
            // Load the workbook using the old password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Assign the new password for saving (this will encrypt the file with the new password)
            workbook.Settings.Password = newPassword;

            // Save the workbook with the new password
            string outputFile = "encrypted_output.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as {outputFile}");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors, such as invalid password
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
