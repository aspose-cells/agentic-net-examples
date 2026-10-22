// Title: Detect Excel file format and prompt for password to decrypt workbook using Aspose.Cells in C#
// Common Searches: how to automatically detect Excel file format and open password‑protected workbook with Aspose.Cells C# | Aspose.Cells load workbook of unknown format and ask user for password if encrypted | C# code to open encrypted Excel file after auto format detection using Aspose.Cells | prompt for workbook password when loading Excel file with LoadOptions Auto in Aspose.Cells | detect and decrypt password‑protected Excel workbook before processing in C#
// Tags: auto format detection LoadOptions Aspose.Cells | load encrypted Excel workbook C# | decrypt workbook using Aspose.Cells | handle unknown Excel format Aspose.Cells | save decrypted workbook C# | user password prompt for workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the source Excel file exists, uses Aspose.Cells LoadOptions with LoadFormat.Auto to automatically determine the workbook format, catches exceptions that indicate password protection, prompts the user for the password, reloads the workbook with the supplied password, lists worksheet names, and finally saves a decrypted copy of the workbook to a specified output path.
class Program
{
    static void Main()
    {
        // Path to the source workbook (any supported format)
        string sourcePath = @"C:\Data\inputFile.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        Workbook workbook = null;

        // Attempt to load the workbook; if password is required, prompt the user
        try
        {
            // 1. Detect file format automatically
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            workbook = new Workbook(sourcePath, loadOptions);
        }
        catch (Exception ex)
        {
            // Check if the exception indicates a password‑protected file
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.Write("Enter password for the workbook: ");
                string password = Console.ReadLine();

                try
                {
                    // Reload the workbook with the supplied password
                    LoadOptions pwdOptions = new LoadOptions(LoadFormat.Auto) { Password = password };
                    workbook = new Workbook(sourcePath, pwdOptions);
                }
                catch (Exception pwdEx)
                {
                    Console.WriteLine($"Failed to load workbook with password: {pwdEx.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }
        }

        // At this point the workbook is loaded (and decrypted if it was protected)
        Console.WriteLine("Worksheets in the workbook:");
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Console.WriteLine("- " + ws.Name);
        }

        // Save the processed workbook (optional – here we save a copy)
        string outputPath = @"C:\Data\outputFile.xlsx";

        try
        {
            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
