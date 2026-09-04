// Title: How to add password protection to an existing Excel workbook and save it with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, applies full workbook protection using a custom password, and saves the protected file to a new path. | Demonstrate using Workbook.Protect with ProtectionType.All in Aspose.Cells to secure an Excel workbook's structure and windows, then persist the changes.
// Common Searches: aspnet protect existing Excel file with password using Aspose.Cells | C# Aspose.Cells Workbook.Protect example for .xlsx | save password‑protected workbook to a different file Aspose.Cells .NET | apply full protection (structure and windows) to Excel workbook programmatically | how to verify input file exists before protecting with Aspose.Cells
// Tags: Aspose.Cells Workbook.Protect password | protect Excel workbook structure C# | save password‑protected .xlsx Aspose.Cells | apply ProtectionType.All Aspose.Cells | load and protect existing workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program checks for the presence of input.xlsx, loads it with Aspose.Cells, applies full workbook protection using the password "MySecurePassword", and saves the protected workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply password protection (protects structure and windows)
            workbook.Protect(ProtectionType.All, "MySecurePassword");

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Protected workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
