// Title: How to unprotect an Excel workbook’s structure and set a new password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, removes any workbook structure protection, then protects the structure with a new password and saves the file. | Show how to confirm that the workbook structure protection password has been changed by applying ProtectionType.Structure with a new password and re‑saving the workbook.
// Common Searches: aspnet change workbook structure password Aspose.Cells | c# remove Excel sheet protection and set new password using Aspose.Cells | how to unprotect and re‑protect Excel workbook structure programmatically in .NET | Aspose.Cells protect workbook structure with custom password example
// Tags: Aspose.Cells Workbook.Unprotect method | Aspose.Cells Workbook.Protect with ProtectionType.Structure | C# update Excel workbook structure password | load and save .xlsx with Aspose.Cells after protection change | verify workbook structure protection in .NET

using System;
using System.IO;
using Aspose.Cells;

// // Loads "input.xlsx", removes existing structure protection, applies new structure protection with password "NewPass123", saves as "output.xlsx", and writes status messages to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string newPassword = "NewPass123";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove any existing structure protection (empty password works if none is set)
            workbook.Unprotect("");

            // Apply new structure protection with the specified password
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Inform the user that protection has been applied
            Console.WriteLine("Structure protection applied.");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
