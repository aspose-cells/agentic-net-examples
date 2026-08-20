// Title: C# – Unprotect Excel workbook structure, rename sheets, and re‑protect with Azure Key Vault password (Aspose.Cells)
// Description: Loads an XLSX file, retrieves a password from Azure Key Vault (or environment), removes workbook‑structure protection, safely renames selected worksheets, reapplies structure protection with the same password, and saves the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | unprotect workbook structure | rename Excel worksheets programmatically | protect workbook with password | Azure Key Vault password retrieval | CellsHelper.CreateSafeSheetName | ProtectionType.Structure | secure Excel file handling | .NET workbook protection example
// Common Searches: Aspose.Cells unprotect workbook and rename sheets C# | protect Excel workbook with Azure Key Vault password .NET | rename worksheets after unprotecting workbook using Aspose.Cells | C# code to change sheet names in a protected Excel file | how to use Azure Key Vault with Aspose.Cells password
// Developer Intent: Remove structure protection from an Excel workbook, rename specific worksheets, then re‑apply protection using a password fetched securely (e.g., from Azure Key Vault) with Aspose.Cells for .NET.
// Use Cases: Automated nightly job that updates sheet names in a protected template before distribution. | Enterprise solution that stores workbook passwords in Azure Key Vault to avoid hard‑coding credentials. | Batch processing pipeline that renames multiple worksheets in a secured workbook as part of data preparation.
// AI Prompts: Write C# code with Aspose.Cells to unprotect a workbook structure using a password from Azure Key Vault, rename the first three sheets, and protect the workbook again. | Show how to replace the placeholder GetWorkbookPassword method with actual Azure Key Vault SDK calls for secure password retrieval. | Explain error handling for cases where the workbook is not protected or the supplied password is invalid when using Aspose.Cells protection APIs.

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX file, retrieves a password from Azure Key Vault (or environment), removes workbook‑structure protection, safely renames selected worksheets, reapplies structure protection with the same password, and saves the result using Aspose.Cells for .NET.
class WorkbookProtectionHandler
{
    static void Main()
    {
        try
        {
            // Retrieve the password (replace with actual retrieval logic if needed)
            string password = GetWorkbookPassword();

            // Verify input file exists before loading
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect workbook structure if it is protected with a password
            if (workbook.IsWorkbookProtectedWithPassword)
            {
                workbook.Unprotect(password);
            }

            // Rename worksheets (example: rename first three sheets)
            string[] newNames = { "Summary", "Data", "Report" };
            for (int i = 0; i < Math.Min(newNames.Length, workbook.Worksheets.Count); i++)
            {
                // Ensure the new name is a valid Excel sheet name
                string safeName = CellsHelper.CreateSafeSheetName(newNames[i]);
                workbook.Worksheets[i].Name = safeName;
            }

            // Re‑protect the workbook structure with the same password
            workbook.Protect(ProtectionType.Structure, password);

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder for password retrieval; replace with Azure Key Vault or other secure source as needed
    private static string GetWorkbookPassword()
    {
        // Try environment variable first, then fallback to a default password
        string envPassword = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD");
        return string.IsNullOrEmpty(envPassword) ? "defaultPassword" : envPassword;
    }
}
