// Title: How to load only the "Summary" and "Data" worksheets from an Excel file using Aspose.Cells LoadFilter in C#
// Common Searches: C# Aspose.Cells load specific worksheets using LoadFilter | Load only selected sheets from an Excel workbook with Aspose.Cells | Aspose.Cells LoadOptions filter to open only Summary and Data sheets | Selective worksheet loading in C# with Aspose.Cells LoadFilter example | Save a workbook after loading only certain worksheets using Aspose.Cells
// Tags: LoadFilter selective worksheet loading | Aspose.Cells load specific sheets C# | Workbook load options filter Excel | Save filtered workbook Aspose.Cells | Custom LoadFilter implementation C#

using Aspose.Cells;
using System;
using System.IO;

// The program defines a custom SheetLoadFilter that overrides ShouldLoadSheet to return true only for worksheets named "Summary" or "Data". A LoadOptions object is created with this filter, then the source Excel file is loaded using those options. The resulting workbook, containing only the two specified sheets, is saved to a new file.
class Program
{
    // Custom filter to load only specific worksheets
    private class SheetLoadFilter : LoadFilter
    {
        // Hide the base method (non‑virtual in some Aspose.Cells versions)
        public new bool ShouldLoadSheet(string sheetName)
        {
            return sheetName == "Summary" || sheetName == "Data";
        }
    }

    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Verify that the source file exists
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new SheetLoadFilter()
            };

            // Load the workbook using the filter
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the filtered workbook (optional)
            string outputPath = "filtered.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
