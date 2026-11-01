// Title: Add a modify password to an existing ODS workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set modify password on existing ODS workbook | protect ODS spreadsheet from editing using Workbook.Protect in .NET | apply write protection to ODS file without changing its data Aspose.Cells | how to add password protection to an ODS workbook programmatically in C# | save ODS workbook with modify password while preserving its content using Aspose.Cells
// Tags: ods workbook modify password aspnet cells | aspnet cells protect workbook write protection | c# set ods file password protection | aspose.cells workbook protect method ods | preserve data while applying ODS password protection

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing ODS file with Aspose.Cells, applies write protection using Workbook.Protect(ProtectionType.All, modifyPassword) to require a password for editing, and saves the workbook back to ODS format, preserving all original data. Includes file‑existence check and basic error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.ods";
        const string outputPath = "output.ods";
        const string modifyPassword = "YourModifyPassword";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply write protection with a password
            workbook.Protect(ProtectionType.All, modifyPassword);

            // Save the workbook, preserving all original data
            workbook.Save(outputPath, SaveFormat.Ods);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
