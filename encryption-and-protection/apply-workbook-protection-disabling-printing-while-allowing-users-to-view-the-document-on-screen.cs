// Title: Password‑protect an Excel workbook and attempt to block printing with Aspose.Cells for .NET
// Common Searches: how to protect an Excel workbook with a password using Aspose.Cells in C# and stop users from printing | Aspose.Cells C# disable print command on a protected workbook | set view‑only protection on an Excel file while preventing printing with Aspose.Cells .NET | C# code example for workbook.Protect to block printing in Aspose.Cells | which Aspose.Cells version supports printing restriction on protected workbooks
// Tags: password protect workbook Aspose.Cells C# | printing restriction workbook protection Aspose.Cells | Aspose.Cells workbook.Protect API usage | protect Excel structure windows Aspose.Cells | view‑only workbook protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file, applies full workbook protection (structure and windows) with a password via workbook.Protect(ProtectionType.All, "MySecretPassword"), notes that disabling printing may not be available in all Aspose.Cells versions, saves the protected file, and includes basic error handling.
class WorkbookProtectionExample
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "protected.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook (structure and windows) with a password
            // ProtectionType.All protects the entire workbook
            workbook.Protect(ProtectionType.All, "MySecretPassword");

            // Note: Disabling printing via WorkbookProtection is not available in all versions.
            // If needed, additional logic can be added here using the appropriate API for the target version.

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook has been protected and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions to prevent the application from crashing
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
