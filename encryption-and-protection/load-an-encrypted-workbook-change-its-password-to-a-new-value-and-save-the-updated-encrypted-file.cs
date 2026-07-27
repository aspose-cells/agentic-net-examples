// Title: C# – Change Password of an Encrypted Excel Workbook with Aspose.Cells
// Description: This example demonstrates how to open a password‑protected Excel file using Aspose.Cells LoadOptions, assign a new password via workbook.Settings.Password (which automatically re‑encrypts the file), and save the workbook under a new name. It includes robust handling for Aspose.Cells‑specific and generic exceptions.
// Keywords: Aspose.Cells | C# | .NET | encrypted Excel workbook | change workbook password | re‑encrypt Excel file | LoadOptions password | Excel file protection | programmatic password reset | Excel encryption .NET
// Common Searches: Aspose.Cells change Excel file password C# | How to re‑encrypt an encrypted .xlsx with a new password using Aspose.Cells | Load password protected workbook and save with new password .NET | Update password of encrypted Excel workbook programmatically | C# code to reset Excel workbook password Aspose
// Developer Intent: Replace the existing password of a password‑protected Excel workbook and save the updated file.
// Use Cases: Migrate legacy spreadsheets to a new security policy by re‑encrypting them with a fresh password. | Automate periodic password rotation for confidential Excel reports in scheduled jobs. | Enable self‑service password reset for users after identity verification without manual file handling.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an encrypted .xlsx with a known password, change the password to a new value, and save the workbook. | Explain how setting workbook.Settings.Password triggers re‑encryption in Aspose.Cells and what other protection settings are preserved. | Provide best‑practice error handling for loading and saving password‑protected workbooks with Aspose.Cells in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to open a password‑protected Excel file using Aspose.Cells LoadOptions, assign a new password via workbook.Settings.Password (which automatically re‑encrypts the file), and save the workbook under a new name. It includes robust handling for Aspose.Cells‑specific and generic exceptions.
class ChangeWorkbookPassword
{
    static void Main()
    {
        // Path to the existing encrypted workbook
        string inputFile = "encrypted.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' not found.");
            return;
        }

        // Current password used to open the workbook
        string oldPassword = "oldpwd";

        // New password to set for the workbook
        string newPassword = "newpwd";

        try
        {
            // Load the workbook with the existing password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = oldPassword
            };
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Update the password (re‑encrypts the file with the new password)
            workbook.Settings.Password = newPassword;

            // Save the workbook with the updated password
            string outputFile = "encrypted_updated.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook password updated successfully. Saved as '{outputFile}'.");
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
