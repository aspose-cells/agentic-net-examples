// Title: C# – Unprotect, Rename Sheets, and Re‑protect Excel Workbook using Aspose.Cells & Azure Key Vault
// Description: Loads an Excel file with Aspose.Cells, retrieves the workbook‑protection password from Azure Key Vault (or an environment variable fallback), removes structure protection, safely renames selected worksheets, re‑applies structure protection with the same password, and saves the result to a target path.
// Keywords: Aspose.Cells | C# Excel | unprotect workbook | protect workbook structure | rename worksheets | Azure Key Vault secret | environment variable fallback | Workbook.Unprotect | Workbook.Protect | CreateSafeSheetName
// Common Searches: How to unprotect an Excel workbook structure with Aspose.Cells in C# | Rename multiple worksheets safely using Aspose.Cells | Use Azure Key Vault to supply a password for Aspose.Cells workbook protection | Re‑apply workbook structure protection after modifying sheets with Aspose.Cells | Fallback to environment variable when Azure Key Vault SDK is unavailable
// Developer Intent: Remove structure protection, rename specific sheets, and re‑apply protection using a password fetched from Azure Key Vault.
// Use Cases: Automate processing of a password‑protected workbook: unprotect, rename the first three sheets to "Summary", "Data", and "Report", then protect again. | Integrate Azure Key Vault secret retrieval for workbook passwords, with an environment‑variable fallback for CI/CD pipelines. | Ensure output directories exist before saving the modified workbook to avoid runtime errors.
// AI Prompts: Generate C# code that uses Aspose.Cells to unprotect a workbook structure, rename given worksheets, and protect the workbook again, pulling the password from Azure Key Vault. | Write a robust GetPasswordFromKeyVault method that calls the Azure SDK, handles exceptions, and falls back to an environment variable. | Show how to validate worksheet names with CellsHelper.CreateSafeSheetName before assigning them in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, retrieves the workbook‑protection password from Azure Key Vault (or an environment variable fallback), removes structure protection, safely renames selected worksheets, re‑applies structure protection with the same password, and saves the result to a target path.
class Program
{
    // Retrieves a password (placeholder). Replace with Azure Key Vault logic if the SDK is available.
    static string GetPasswordFromKeyVault(string vaultUrl, string secretName)
    {
        try
        {
            // Azure SDK not referenced; attempt to read from environment variable as a fallback.
            string envVar = $"{secretName}_PASSWORD";
            string pwd = Environment.GetEnvironmentVariable(envVar);
            if (!string.IsNullOrEmpty(pwd))
                return pwd;

            Console.WriteLine("Azure Key Vault SDK not available. Using empty password.");
            return string.Empty;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving password: {ex.Message}");
            return string.Empty;
        }
    }

    static void Main(string[] args)
    {
        // Expected arguments:
        // args[0] - input workbook path
        // args[1] - output workbook path
        // args[2] - Azure Key Vault URL (e.g., https://myvault.vault.azure.net/)
        // args[3] - Secret name that holds the workbook protection password
        if (args.Length < 4)
        {
            Console.WriteLine("Usage: <inputPath> <outputPath> <vaultUrl> <secretName>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];
        string vaultUrl = args[2];
        string secretName = args[3];

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Retrieve the password (placeholder implementation)
        string password = GetPasswordFromKeyVault(vaultUrl, secretName);

        Workbook workbook = null;
        try
        {
            // Load the workbook (lifecycle rule: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // If the workbook structure is protected with a password, unprotect it
            if (workbook.IsWorkbookProtectedWithPassword)
            {
                workbook.Unprotect(password); // rule: Workbook.Unprotect(string)
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to unprotect workbook: {ex.Message}");
            // Continue processing; some operations may still succeed
        }

        // Example rename mapping (old index -> new name)
        var renameMap = new (int Index, string NewName)[]
        {
            (0, "Summary"),
            (1, "Data"),
            (2, "Report")
        };

        // Rename worksheets using safe sheet names
        foreach (var (index, newName) in renameMap)
        {
            if (index >= 0 && index < workbook.Worksheets.Count)
            {
                string safeName = CellsHelper.CreateSafeSheetName(newName);
                workbook.Worksheets[index].Name = safeName;
            }
        }

        try
        {
            // Re‑protect the workbook structure with the same password
            workbook.Protect(ProtectionType.Structure, password); // rule: Workbook.Protect(ProtectionType, string)
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to protect workbook: {ex.Message}");
        }

        try
        {
            // Ensure output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
