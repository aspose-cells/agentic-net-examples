// Title: Upgrade the encryption algorithm of a password‑protected Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: asp.net load encrypted xlsx and save with stronger encryption using Aspose.Cells | c# upgrade encryption algorithm of a password‑protected Excel workbook with Aspose.Cells | re‑encrypt an existing encrypted workbook to default strong encryption Aspose.Cells .NET | how to change encryption level of an XLSX file in C# with Aspose.Cells | load workbook with password and save to apply new encryption settings Aspose.Cells
// Tags: load encrypted workbook Aspose.Cells C# | upgrade Excel encryption algorithm Aspose.Cells | save workbook with default strong encryption Aspose.Cells | re‑encrypt XLSX file using Aspose.Cells .NET | change encryption level Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the encrypted Excel file, loads it with the original password via LoadOptions, retains the password in the workbook settings, and saves the workbook. Saving automatically applies Aspose.Cells' default stronger encryption algorithm, effectively upgrading the workbook's protection.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the source and destination workbooks
            string inputPath = "EncryptedInput.xlsx";
            string outputPath = "UpgradedEncryptedOutput.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Password used to open the encrypted workbook
            string password = "oldPassword";

            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Preserve the existing password (encryption algorithm will use the default strong algorithm)
            workbook.Settings.Password = password;

            // Save the workbook; the new encryption settings will be applied
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
