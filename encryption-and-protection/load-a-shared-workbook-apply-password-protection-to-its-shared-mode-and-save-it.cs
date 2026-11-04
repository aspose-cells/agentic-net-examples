// Title: Load a shared Excel workbook, apply full password protection, and save it using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to protect a shared Excel workbook with a password | how to use Workbook.Protect with ProtectionType.All in Aspose.Cells | save an encrypted Excel file after loading it with Aspose.Cells .NET | apply password protection to all aspects of a workbook using Aspose.Cells C# | protect shared mode of an existing .xlsx file with Aspose.Cells API
// Tags: Aspose.Cells workbook password protection C# | Workbook.Protect all features Aspose.Cells | encrypt shared Excel workbook Aspose.Cells | save password‑protected .xlsx with Aspose.Cells | load existing workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the source file, loads the shared workbook into an Aspose.Cells Workbook object, applies full password protection with Workbook.Protect(ProtectionType.All, password), and saves the protected file to a new location, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "SharedWorkbook.xlsx";
            const string outputPath = "SharedWorkbook_Protected.xlsx";
            const string password = "MySecretPassword";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply password protection to the workbook (protect all aspects)
            workbook.Protect(ProtectionType.All, password);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
