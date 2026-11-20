// Title: Unprotect an Excel workbook structure, rename specific sheets, and re‑protect it with a password retrieved from Azure Key Vault using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unprotect workbook structure with password from Azure Key Vault | Rename multiple worksheets after unprotecting an Excel file using Aspose.Cells in .NET | How to protect Excel workbook structure with a password using Aspose.Cells C# | Load, modify, and save an XLSX file while preserving protection with Aspose.Cells | Retrieve Excel password from Azure Key Vault for Aspose.Cells workbook operations
// Tags: Aspose.Cells unprotect workbook structure C# | Aspose.Cells rename worksheets C# | Aspose.Cells protect workbook structure password C# | Azure Key Vault password retrieval for Excel Aspose.Cells | Aspose.Cells load and save XLSX C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook, obtains the protection password (placeholder for Azure Key Vault retrieval), calls Workbook.Unprotect with that password, renames the first three worksheets to "SalesData", "Summary", and "Archive", then re‑protects the workbook structure using Workbook.Protect with the same password, and finally saves the modified file, handling I/O and protection errors.
class WorkbookProcessor
{
    static void Main()
    {
        try
        {
            // Retrieve workbook password (replace with actual retrieval logic, e.g., Azure Key Vault)
            string password = GetWorkbookPassword();

            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Unprotect the workbook structure using the retrieved password
            try
            {
                // Workbook.Unprotect accepts only the password
                workbook.Unprotect(password);
            }
            catch (Exception unprotectEx)
            {
                Console.WriteLine($"Failed to unprotect workbook structure: {unprotectEx.Message}");
                // Continue if the workbook was not protected
            }

            // Rename worksheets as required
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets[0].Name = "SalesData";
            if (workbook.Worksheets.Count > 1)
                workbook.Worksheets[1].Name = "Summary";
            if (workbook.Worksheets.Count > 2)
                workbook.Worksheets[2].Name = "Archive";

            // Re‑protect the workbook structure with the same password
            try
            {
                workbook.Protect(ProtectionType.Structure, password);
            }
            catch (Exception protectEx)
            {
                Console.WriteLine($"Failed to protect workbook structure: {protectEx.Message}");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder for password retrieval logic
    private static string GetWorkbookPassword()
    {
        // Example: read from an environment variable; replace with your own method as needed
        return Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD") ?? string.Empty;
    }
}
