// Title: Check if an Excel .xlsx workbook requires an opening password or a modify password using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells and prints whether an open password is set. | Demonstrate how to query the write‑protection flag of a workbook via Aspose.Cells and output the modify‑password requirement.
// Common Searches: Aspose.Cells C# determine if an Excel file is password‑protected for opening | How to detect write‑protected Excel workbook using Aspose.Cells .NET | C# code to read Workbook.Settings.Password and Workbook.Settings.WriteProtection.IsWriteProtected | Check both open and modify password status of .xlsx with Aspose.Cells | Retrieve Excel workbook protection flags without providing a password in Aspose.Cells
// Tags: Workbook.Settings.Password presence check Aspose.Cells | Workbook.Settings.WriteProtection flag evaluation Aspose.Cells | Aspose.Cells Excel password protection detection .NET | Aspose.Cells modify protection status query | C# log workbook protection flags

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel file with Aspose.Cells, checks Workbook.Settings.Password to see if an opening password is required, checks Workbook.Settings.WriteProtection.IsWriteProtected for modify protection, and writes both statuses to the console.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "sample.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook without providing any passwords
            Workbook workbook = new Workbook(filePath);

            // Determine if a password is required to open the workbook
            bool requiresOpenPassword = !string.IsNullOrEmpty(workbook.Settings.Password);

            // Determine if a password is required to modify (write‑protect) the workbook
            bool requiresModifyPassword = workbook.Settings.WriteProtection.IsWriteProtected;

            // Log the statuses
            Console.WriteLine($"Requires password to open: {requiresOpenPassword}");
            Console.WriteLine($"Requires password to modify: {requiresModifyPassword}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., corrupted file, unsupported format)
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
