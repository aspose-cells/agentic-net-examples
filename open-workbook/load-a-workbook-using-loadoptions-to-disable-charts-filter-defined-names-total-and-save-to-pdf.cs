// Title: Load an Excel workbook without charts, keep only the "Total" defined name, and save as PDF using Aspose.Cells for .NET
// Description: A C# example that verifies the source file, creates a LoadFilter to exclude chart data, loads the workbook with LoadOptions, removes all defined names except "Total", and exports the result to a PDF file.
// Keywords: Aspose.Cells LoadOptions chart exclusion | C# load Excel without charts | remove named ranges Aspose.Cells | save workbook as PDF .NET | LoadFilter defined names | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells load Excel without charts | how to keep only one defined name when converting to PDF | exclude chart objects using LoadFilter Aspose.Cells | remove all named ranges except specific one C# | convert Excel to PDF with Aspose.Cells while stripping charts
// Developer Intent: Load an Excel file, skip chart objects, retain only the "Total" named range, and generate a PDF.
// Use Cases: Create lightweight PDF reports by omitting chart graphics. | Prepare data‑only workbooks for downstream processing by preserving a single named range. | Reduce PDF file size by eliminating unnecessary named ranges and chart data.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions with a LoadFilter to load an Excel workbook while ignoring charts and then saves it as a PDF. | Show how to iterate the workbook's Names collection and delete every defined name except "Total" before exporting to PDF. | Explain the impact of excluding chart data via LoadOptions on PDF conversion performance and file size.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A C# example that verifies the source file, creates a LoadFilter to exclude chart data, loads the workbook with LoadOptions, removes all defined names except "Total", and exports the result to a PDF file.
class LoadWorkbookAndSavePdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Create LoadOptions with a custom LoadFilter that excludes charts.
            LoadFilter loadFilter = new LoadFilter(LoadDataFilterOptions.All ^ LoadDataFilterOptions.Chart);
            LoadOptions loadOptions = new LoadOptions { LoadFilter = loadFilter };

            // Load the workbook using the specified LoadOptions.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // ------------------------------------------------------------
            // Filter defined names: keep only the name "Total".
            // Defined names are stored in the workbook's Names collection.
            // ------------------------------------------------------------
            NameCollection names = workbook.Worksheets.Names;
            for (int i = names.Count - 1; i >= 0; i--)
            {
                // Each item is an Aspose.Cells.Name object.
                Aspose.Cells.Name definedName = (Aspose.Cells.Name)names[i];
                // The name string is accessed via the Text property.
                if (!string.Equals(definedName.Text, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    names.RemoveAt(i);
                }
            }

            // Save the workbook as PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
