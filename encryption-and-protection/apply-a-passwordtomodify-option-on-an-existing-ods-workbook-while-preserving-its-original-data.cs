// Title: How to add a modify‑password protection to an existing ODS workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set modify password for existing ODS workbook | How to protect an ODS file with a password to edit using Aspose.Cells .NET | Apply password‑to‑modify protection without changing data in ODS spreadsheet Aspose.Cells | Workbook.Protect example for ODS files in C# | Saving ODS workbook with password protection using Aspose.Cells
// Tags: protect ODS workbook with password Aspose.Cells | set modify password ODS Aspose.Cells .NET | Workbook.Protect method ODS format | save protected ODS file Aspose.Cells | apply password protection to existing ODS spreadsheet

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing ODS workbook, applies a password‑to‑modify protection using Workbook.Protect with ProtectionType.All, and saves the file, preserving all original data.
class Program
{
    static void Main()
    {
        const string inputPath = "input.ods";
        const string outputPath = "output.ods";
        const string password = "YourPasswordHere";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook so that a password is required to modify it
            workbook.Protect(ProtectionType.All, password);

            // Save the workbook, preserving all original data
            workbook.Save(outputPath, SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
