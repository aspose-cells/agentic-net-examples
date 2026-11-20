// Title: How to protect every worksheet in an Excel file with a password using Aspose.Cells for .NET while keeping original formatting
// Common Searches: C# Aspose.Cells protect all worksheets with the same password | Apply worksheet protection to each sheet in a workbook using Aspose.Cells .NET | Set password for every worksheet in an Excel workbook programmatically with Aspose.Cells | Iterate through worksheets and protect them without losing formatting in C# | Batch protect Excel sheets using Aspose.Cells Protect method
// Tags: Aspose.Cells worksheet password protection | batch sheet protection C# Aspose.Cells | preserve formatting while protecting Excel sheets | ProtectionType.All Aspose.Cells example | iterate worksheets protect password .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, loops through all worksheets, applies full protection with a specified password using ProtectionType.All, and saves the workbook, ensuring the original formatting remains unchanged.
class WorksheetProtectionExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "SecurePassword";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Protect each worksheet with a password
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Apply protection to all aspects of the sheet
                sheet.Protect(ProtectionType.All);
                // Set the password for the protection
                sheet.Protection.Password = password;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
