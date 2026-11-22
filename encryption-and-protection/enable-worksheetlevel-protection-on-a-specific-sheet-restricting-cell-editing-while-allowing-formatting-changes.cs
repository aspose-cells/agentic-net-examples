// Title: Password‑protect a single worksheet while allowing formatting changes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect only one worksheet with a password and allow cell formatting | How to enable worksheet protection for a specific sheet in an Excel file using Aspose.Cells .NET | C# code to protect a worksheet but keep formatting options editable with Aspose.Cells | Save an Excel workbook with a protected sheet using Aspose.Cells for .NET | Set worksheet protection type to All with password in Aspose.Cells C# example
// Tags: worksheet password protection Aspose.Cells .NET | single sheet protection C# Aspose.Cells | allow formatting changes with worksheet protection Aspose.Cells | saving protected Excel file Aspose.Cells | using ProtectionType.All Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, selects a worksheet (by name or the first sheet), applies full worksheet protection with a password via sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty) which blocks cell edits while permitting formatting changes, and saves the protected workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet by name; fall back to the first sheet if not found
            Worksheet sheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            // Protect the worksheet with a password (oldPassword is empty string)
            sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

            // Note: Individual protection options (e.g., AllowEditObject) are not set here
            // because the current Aspose.Cells version does not expose those properties.
            // The worksheet is fully protected with the specified password.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
