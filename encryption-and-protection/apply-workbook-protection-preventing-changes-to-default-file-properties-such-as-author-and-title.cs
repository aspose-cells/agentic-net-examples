// Title: How to password‑protect an entire Excel workbook (structure and windows) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect workbook structure and windows with password | C# code to apply ProtectionType.All on an Excel file using Aspose.Cells | Save a password‑protected Excel workbook with Aspose.Cells .NET | Prevent editing of workbook properties in Excel via Aspose.Cells protection | Apply full workbook protection before saving with Aspose.Cells
// Tags: Aspose.Cells workbook full protection C# | ProtectionType.All usage Aspose.Cells | password protected Excel workbook .NET | protect workbook structure and windows Aspose.Cells | save protected workbook with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, applies full protection (structure and windows) with a password using workbook.Protect(ProtectionType.All, password), and saves the protected file as output.xlsx while handling any exceptions.
class WorkbookProtectionExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "StrongPassword123";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the entire workbook (structure and windows) with a password
            workbook.Protect(ProtectionType.All, password);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
