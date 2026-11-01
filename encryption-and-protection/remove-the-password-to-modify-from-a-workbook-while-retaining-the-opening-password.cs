// Title: Remove the modify (edit) password from an Excel workbook while preserving the opening password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET remove edit password from existing XLSX file | keep opening password but delete modify protection in Excel using C# | load password-protected workbook and save without modify password Aspose.Cells | C# code to clear workbook modify password while retaining read password | Aspose.Cells remove workbook protection for editing but keep read-only password
// Tags: remove workbook modify password Aspose.Cells | retain workbook opening password .NET | load encrypted XLSX with password Aspose.Cells | save Excel file without edit protection C# | Aspose.Cells workbook.Settings.Password usage

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX file with its opening (read) password via LoadOptions, reassigns the same password to workbook.Settings.Password to keep the file protected for opening, and saves the workbook without specifying a modify password, thereby removing edit‑level protection while preserving the original opening password.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string openPassword = "OpenPassword123";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The file '{inputPath}' was not found.");

            // Load the workbook with the opening (read) password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = openPassword
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Preserve the opening password for the saved file
            workbook.Settings.Password = openPassword;

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
