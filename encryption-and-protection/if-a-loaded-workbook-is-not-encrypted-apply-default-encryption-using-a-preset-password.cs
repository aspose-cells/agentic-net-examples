// Title: Apply default password encryption to an unprotected Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, assign a preset password to the workbook, and save it as an encrypted file in C#. | Verify the input file exists, create the output folder if necessary, then set workbook.Settings.Password to protect the workbook before saving. | Overwrite a previously saved workbook with password protection using Aspose.Cells, ensuring the result is saved in encrypted XLSX format.
// Common Searches: Aspose.Cells C# set default password for an existing workbook without existing protection | How to encrypt an already saved Excel file with a predefined password using Aspose.Cells .NET | Save XLSX with password protection only when file is not encrypted in C# | Create output directory and apply workbook.Settings.Password before saving with Aspose.Cells
// Tags: default password encryption Aspose.Cells | workbook.Settings.Password C# | save encrypted XLSX Aspose.Cells | ensure output directory before saving Aspose.Cells | apply password to unprotected workbook .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file, checks that the file exists, assigns a preset password via workbook.Settings.Password, creates the output directory if needed, and saves the workbook as an encrypted XLSX file using Aspose.Cells for .NET.
class WorkbookEncryption
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";
        // Path for the encrypted output (can overwrite the original)
        string outputPath = "input_encrypted.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply encryption with a preset password (overwrites any existing password)
            string presetPassword = "MyDefaultPassword123";
            workbook.Settings.Password = presetPassword;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (encrypted)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
