// Title: Unprotect Excel workbook structure, rename worksheets, and re‑protect using a password from Azure Key Vault with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells unprotect workbook structure and rename worksheets | How to retrieve an Excel password from Azure Key Vault and protect a workbook with Aspose.Cells | Rename Excel sheets after unprotecting a workbook using Aspose.Cells .NET | Protect workbook structure with a password using Aspose.Cells and Azure Key Vault | Load an Excel file, change sheet names, and re‑apply protection in C# Aspose.Cells
// Tags: unprotect workbook structure Aspose.Cells C# | rename worksheets Aspose.Cells API | protect workbook structure with password Aspose.Cells | Azure Key Vault integration for Excel password C# | Aspose.Cells workbook protection workflow .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, obtains a password (placeholder for Azure Key Vault retrieval), unprotects the workbook's structure, renames specific worksheets, re‑applies structure protection using the same password, and saves the modified file to a new location.
class WorkbookProcessor
{
    static void Main()
    {
        // Parameters (replace with actual values or retrieve from configuration)
        string inputPath = @"C:\Docs\Workbook.xlsx";
        string outputPath = @"C:\Docs\Workbook_Protected.xlsx";

        // Retrieve the password (replace with Azure Key Vault logic if available)
        string password = GetWorkbookPassword();

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (structure and windows) using the retrieved password
            workbook.Unprotect(password);

            // Rename sheets as required
            Worksheet sheet1 = workbook.Worksheets["Sheet1"];
            if (sheet1 != null)
                sheet1.Name = "Data";

            Worksheet sheet2 = workbook.Worksheets["Sheet2"];
            if (sheet2 != null)
                sheet2.Name = "Summary";

            // Re‑protect the workbook structure with the same password
            workbook.Protect(ProtectionType.Structure, password);

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }

    // Retrieves the workbook password; replace with Azure Key Vault retrieval if needed
    private static string GetWorkbookPassword()
    {
        // Attempt to read from an environment variable; fallback to a default (empty) password
        string envPassword = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD");
        return string.IsNullOrEmpty(envPassword) ? string.Empty : envPassword;
    }
}
