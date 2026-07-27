// Title: C# – Load Excel with Aspose.Cells, exclude charts, keep only the “Total” named range, and export to PDF
// Description: The sample verifies the source file, creates a LoadFilter that omits chart objects, applies it via LoadOptions, removes every defined name except "Total" from the Worksheets.Names collection, and saves the resulting workbook as a PDF. This approach reduces memory usage and file size when converting Excel to PDF.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter | exclude charts | C# Excel to PDF | remove named ranges | keep defined name | Total named range | .NET | PDF conversion | performance optimization
// Common Searches: Aspose.Cells load workbook without charts C# | How to keep only a specific defined name in Aspose.Cells | Export Excel to PDF while ignoring charts Aspose.Cells | LoadFilter chart exclusion example .NET | C# remove all named ranges except one before PDF conversion
// Developer Intent: Load an Excel file, skip chart objects, retain only the "Total" defined name, and save the workbook as a PDF using Aspose.Cells.
// Use Cases: Generate lightweight PDF reports that contain only the data referenced by the "Total" named range, eliminating chart overhead. | Automate batch conversion of financial workbooks where only the summary range is needed, improving conversion speed and reducing output size. | Create archival PDFs for regulatory compliance by stripping charts and unnecessary named ranges from source spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells to open an Excel file, use LoadFilter to exclude charts, keep only the defined name "Total", and save the result as a PDF. | Show how to iterate through workbook.Worksheets.Names and delete all names except a specified one before exporting to PDF using Aspose.Cells. | Explain how configuring LoadOptions with a chart‑exclusion LoadFilter improves performance when converting Excel to PDF in .NET.

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the source file, creates a LoadFilter that omits chart objects, applies it via LoadOptions, removes every defined name except "Total" from the Worksheets.Names collection, and saves the resulting workbook as a PDF. This approach reduces memory usage and file size when converting Excel to PDF.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string sourcePath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
                return;
            }

            // Load filter to exclude charts while loading the workbook
            LoadFilter loadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);
            LoadOptions loadOptions = new LoadOptions { LoadFilter = loadFilter };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Keep only the defined name "Total" and remove all others
            // Defined names are stored in the Worksheets.Names collection
            for (int i = workbook.Worksheets.Names.Count - 1; i >= 0; i--)
            {
                Name definedName = workbook.Worksheets.Names[i];
                // The Name object's identifier is accessed via the Text property
                if (!string.Equals(definedName.Text, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    workbook.Worksheets.Names.RemoveAt(i);
                }
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
