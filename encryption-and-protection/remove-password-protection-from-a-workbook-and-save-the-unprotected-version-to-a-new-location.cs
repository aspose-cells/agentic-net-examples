// Title: Remove Excel Workbook Password with Aspose.Cells for .NET and Save an Unprotected Copy
// Description: Demonstrates how to open a password‑protected XLSX file using Aspose.Cells LoadOptions.Password in C#, then save it to a new location without encryption, producing an unprotected workbook.
// Keywords: Aspose.Cells | remove workbook password | C# load encrypted Excel | LoadOptions.Password | save unencrypted workbook | Excel decryption .NET | Aspose.Cells example
// Common Searches: aspocells open encrypted xlsx with password | remove password from Excel file using C# | save workbook without encryption aspocells | loadoptions.password example aspocells | how to decrypt protected Excel programmatically
// Developer Intent: Open a protected Excel file and create a new version that has no password.
// Use Cases: Automate bulk removal of passwords from Excel reports before feeding them into data pipelines. | Prepare workbooks for cloud storage services that reject encrypted files. | Validate file accessibility after stripping protection to ensure downstream processing can proceed.
// AI Prompts: Write C# code that uses Aspose.Cells to open an encrypted .xlsx with a known password and saves it without any protection. | Explain best practices for handling FileNotFoundException and CellsException when removing workbook encryption. | Provide a script that iterates over a directory of password‑protected Excel files, removes each password, and writes the unprotected files to a target folder.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to open a password‑protected XLSX file using Aspose.Cells LoadOptions.Password in C#, then save it to a new location without encryption, producing an unprotected workbook.
class RemoveWorkbookPassword
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected_workbook.xlsx";

        // Path where the unprotected workbook will be saved
        string outputPath = "unprotected_workbook.xlsx";

        // The password that protects the workbook (file encryption password)
        string password = "mySecretPwd";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook without a password (no encryption)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without password to: {outputPath}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
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
