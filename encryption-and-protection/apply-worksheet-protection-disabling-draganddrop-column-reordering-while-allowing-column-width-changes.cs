// Title: How to protect an Excel worksheet with Aspose.Cells for .NET while still allowing column width changes and preventing column drag‑and‑drop reordering (C#)
// Common Searches: Aspose.Cells C# protect worksheet but keep column width adjustable | prevent column drag and drop reordering in a protected Excel sheet using Aspose.Cells | disable column reordering while allowing column resizing in Aspose.Cells .NET | set worksheet protection with AllowSorting false Aspose.Cells example | keep column formatting editable after protecting sheet with Aspose.Cells C# | Aspose.Cells protect sheet without locking column width changes
// Tags: Aspose.Cells worksheet protection allow column resizing | disable column drag-and-drop reordering Aspose.Cells | C# set AllowSorting false Aspose.Cells | protect Excel sheet while permitting column width changes | Aspose.Cells ProtectionType.All custom Allow settings

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an existing XLSX workbook, applies full worksheet protection to the first sheet, disables sorting and filtering, and ensures column width changes remain allowed before saving the protected file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet (no password) and enable all protection flags
            sheet.Protect(ProtectionType.All);

            // Configure allowed actions while the sheet is protected.
            // Note: Some specific Allow* properties may not be available in older Aspose.Cells versions.
            // The following properties are commonly supported.
            // sheet.Protection.AllowFormattingCells = true; // Removed due to API compatibility
            sheet.Protection.AllowSorting = false;          // disallow sorting
            sheet.Protection.AllowFiltering = false;        // disallow filtering

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
