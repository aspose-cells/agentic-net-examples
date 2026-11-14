// Title: Change the password of an encrypted Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to change the password of an encrypted Excel workbook with Aspose.Cells C# | replace old workbook password with a new one using Aspose.Cells .NET | load a password‑protected XLSX and save it with a different password in C# | programmatically update Excel file encryption password with Aspose.Cells | Aspose.Cells change workbook password without opening the file in Excel
// Tags: Aspose.Cells change workbook password | LoadOptions password property | Workbook.Settings.Password assignment | C# update encrypted XLSX password | save workbook with new password Aspose.Cells | encrypted Excel file password rotation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an encrypted XLSX file using the old password via LoadOptions, assigns a stronger password to Workbook.Settings.Password, and saves the workbook as a new file with the updated encryption.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "EncryptedWorkbook.xlsx";
            const string outputPath = "EncryptedWorkbook_Updated.xlsx";

            // Verify that the source file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the encrypted workbook using the old password
            var loadOptions = new LoadOptions
            {
                Password = "oldPassword123"
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Set a new password for the workbook when saving
            workbook.Settings.Password = "NewStrongerPassword!@#2026";

            // Save the workbook with the new password
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
