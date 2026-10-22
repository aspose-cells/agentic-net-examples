// Title: Change the password of an encrypted Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load password protected XLSX and set a new password | How to replace workbook encryption password with Aspose.Cells in .NET | Update password of an encrypted Excel file using Aspose.Cells example | Programmatically change Excel file password with Aspose.Cells LoadOptions | Save an encrypted workbook with a new password using Aspose.Cells for .NET
// Tags: load encrypted xlsx with LoadOptions Aspose.Cells | set new workbook password Aspose.Cells | save workbook with updated encryption C# | Workbook.Settings.Password property | Aspose.Cells password protection API

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a password‑protected XLSX workbook using LoadOptions, assigns a new password via Workbook.Settings.Password, and saves the file as a newly encrypted workbook.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string inputPath = "EncryptedWorkbook.xlsx";

        // Current password of the workbook
        string oldPassword = "oldPassword";

        // New password to set
        string newPassword = "newPassword";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the encrypted workbook using LoadOptions with the existing password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the new password for the workbook
            workbook.Settings.Password = newPassword;

            // Save the workbook with the new password
            string outputPath = "EncryptedWorkbook_Updated.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
