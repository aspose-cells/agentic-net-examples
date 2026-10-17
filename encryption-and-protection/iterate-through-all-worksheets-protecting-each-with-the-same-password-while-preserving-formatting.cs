// Title: Apply the same password to protect every worksheet in an XLSX workbook while preserving formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect all worksheets with a single password and keep formatting | How to set worksheet protection for every sheet in a workbook using Aspose.Cells .NET | C# code to apply the same password to each worksheet in an Excel file with Aspose.Cells | Bulk protect Excel worksheets while retaining cell styles using Aspose.Cells for .NET | Iterate through worksheets and protect them without losing formatting in Aspose.Cells
// Tags: worksheet protection password Aspose.Cells | protect all worksheets Aspose.Cells .NET | preserve formatting during sheet protection | bulk worksheet protection C# | apply same password to multiple sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, checks its existence, iterates over every worksheet in the workbook, and applies full protection with a single password while retaining all existing formatting. It ensures the output directory exists, saves the protected workbook to a new file, and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string password = "YourPassword";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Protect each worksheet with the specified password while preserving formatting
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Protect(sheet) overload expects (ProtectionType, password, oldPassword)
                sheet.Protect(ProtectionType.All, password, string.Empty);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
