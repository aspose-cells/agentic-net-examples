// Title: How to re‑encrypt an encrypted Excel workbook after editing with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change password of a protected XLSX after modifying cells | load password‑protected workbook, edit data, and save with a new password using Aspose.Cells | automatic re‑encryption of Excel file after changes in .NET | C# example for updating an encrypted Excel file and setting a new workbook password | handle missing encrypted Excel file when using Aspose.Cells LoadOptions
// Tags: load encrypted XLSX with password Aspose.Cells | modify worksheet cell Aspose.Cells C# | assign new workbook password Aspose.Cells | save workbook with encryption Aspose.Cells .NET | automatic re‑encryption after workbook changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads a password‑protected XLSX using LoadOptions, updates cell A1, assigns a new workbook password via Settings.Password, and saves the file with encryption while handling missing file errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "EncryptedInput.xlsx";
            const string outputPath = "EncryptedOutput.xlsx";
            const string oldPassword = "oldPassword";
            const string newPassword = "newPassword";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the encrypted workbook using the original password
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = oldPassword
                };
                var workbook = new Workbook(inputPath, loadOptions);

                // Example modification: change the value of cell A1 in the first worksheet
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Modified content");

                // Set a new password for the workbook before saving
                workbook.Settings.Password = newPassword;

                // Save the workbook with the new password
                var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
