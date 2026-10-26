// Title: Decrypt a password‑protected XLSX workbook, edit cell values, re‑encrypt with a new password, and save as ODS using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load encrypted XLSX with password and modify cells | how to change workbook password and export to ODS with Aspose.Cells .NET | decrypt Excel file, update A1 and B2, re‑encrypt and save as ODS in C# | Aspose.Cells example for opening protected workbook and saving with new password | programmatically edit cells in a password‑protected Excel file and convert to ODS
// Tags: encrypted xlsx loading Aspose.Cells | cell value update Aspose.Cells C# | workbook password change Aspose.Cells | ods saving Aspose.Cells .NET | excel to ods conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the encrypted XLSX file, opens it with the original password via LoadOptions, updates cells A1 and B2, assigns a new password to the workbook, and saves the result as an ODS file, handling Aspose.Cells‑specific and generic exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted source workbook
        string sourcePath = "encrypted_input.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        // Password used to open the existing encrypted workbook
        string oldPassword = "oldPassword";

        // Password to be applied to the workbook after modification
        string newPassword = "newPassword";

        try
        {
            // Load the encrypted workbook using LoadOptions with the old password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // ----- Modify cell values -----
            // Example: set A1 to "Hello", B2 to 123
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B2"].PutValue(123);

            // ----- Re‑encrypt with a new password -----
            // Set the password that will be required to open the saved file
            workbook.Settings.Password = newPassword;

            // Save the workbook as ODS (OpenDocument Spreadsheet) with the new password
            string outputPath = "modified_output.ods";
            workbook.Save(outputPath, SaveFormat.ODS);

            Console.WriteLine("Workbook decrypted, modified, re‑encrypted, and saved as ODS successfully.");
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
