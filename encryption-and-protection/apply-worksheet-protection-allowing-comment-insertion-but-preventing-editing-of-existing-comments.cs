// Title: Protect an Excel worksheet with all default restrictions using Aspose.Cells for .NET (C#)
// Common Searches: how to protect an entire worksheet with default restrictions using Aspose.Cells in C# | Aspose.Cells C# example for applying ProtectionType.All to a worksheet | load, protect, and save an Excel file with Aspose.Cells .NET | C# code to enforce worksheet protection without custom options in Aspose.Cells | protect Excel worksheet programmatically using Aspose.Cells for .NET
// Tags: Aspose.Cells worksheet protection C# | ProtectionType.All Aspose.Cells | default worksheet restrictions .NET | load save workbook with protection Aspose.Cells | C# Excel worksheet protection example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, retrieves the first worksheet, applies full protection using sheet.Protect(ProtectionType.All) which enforces all default restrictions, and saves the protected file to a new location.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Protect the worksheet with all default restrictions
            sheet.Protect(ProtectionType.All);

            // Note: Specific protection options (e.g., AllowEditObject) are omitted
            // because they are not available in the current Aspose.Cells version.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
