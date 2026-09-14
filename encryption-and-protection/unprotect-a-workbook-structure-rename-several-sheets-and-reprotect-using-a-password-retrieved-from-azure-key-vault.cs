// Title: Unprotect an Excel workbook, rename selected worksheets, and re‑protect the structure using a password retrieved from Azure Key Vault with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, obtains a password from Azure Key Vault, calls Workbook.Unprotect, renames the first three worksheets to custom names, and then calls Workbook.Protect with ProtectionType.Structure using the same password. | Create a reusable C# method that fetches a secret from Azure Key Vault (or falls back to an environment variable) and applies it to protect an Aspose.Cells workbook after sheet‑name modifications.
// Common Searches: how to unprotect an Excel workbook with Aspose.Cells and Azure Key Vault password in C# | rename specific worksheets after unprotecting a protected .xlsx using Aspose.Cells .NET | re‑protect workbook structure with the same password retrieved from Azure Key Vault using Aspose.Cells | C# example loading, modifying, and saving a protected Excel file with Aspose.Cells
// Tags: Aspose.Cells workbook protection workflow | Aspose.Cells rename worksheet programmatically | Aspose.Cells protect workbook structure C# | Azure Key Vault secret for Excel password .NET | load modify save .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an .xlsx file, retrieves a password (from Azure Key Vault or an environment variable), optionally unprotects the workbook, renames the first three worksheets to "SalesData", "Summary", and "Archive", re‑protects the workbook structure with the same password, and saves the result to a new file.
class WorkbookProcessor
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = @"C:\Temp\MyWorkbook.xlsx";
        string outputPath = @"C:\Temp\MyWorkbook_Protected.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Retrieve the password (replace with Azure Key Vault logic if needed)
            string password = GetWorkbookPassword();

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                throw new IOException($"Failed to load workbook from '{inputPath}'.", loadEx);
            }

            // Unprotect the workbook (if it is protected)
            try
            {
                workbook.Unprotect(password);
            }
            catch (Exception unprotectEx)
            {
                Console.Error.WriteLine($"Warning: Unable to unprotect workbook with provided password. {unprotectEx.Message}");
            }

            // Rename worksheets as required
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets[0].Name = "SalesData";
            if (workbook.Worksheets.Count > 1)
                workbook.Worksheets[1].Name = "Summary";
            if (workbook.Worksheets.Count > 2)
                workbook.Worksheets[2].Name = "Archive";

            // Re‑protect the workbook structure using the retrieved password
            try
            {
                workbook.Protect(ProtectionType.Structure, password);
            }
            catch (Exception protectEx)
            {
                Console.Error.WriteLine($"Error protecting workbook: {protectEx.Message}");
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }

    // Placeholder for password retrieval.
    // Replace this method with Azure Key Vault integration if required.
    private static string GetWorkbookPassword()
    {
        // Attempt to read from an environment variable; fallback to a default value.
        string? pwd = Environment.GetEnvironmentVariable("WorkbookPassword");
        return string.IsNullOrEmpty(pwd) ? "DefaultPassword123!" : pwd;
    }
}
