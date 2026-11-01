// Title: Convert an Excel workbook to CSV while excluding hyperlinks with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to CSV without including hyperlink URLs | How to ignore hyperlinks when saving a workbook as CSV using Aspose.Cells | C# code to convert .xlsx to .csv and strip hyperlink data with Aspose.Cells | Save workbook as CSV with TxtSaveOptions while omitting hyperlink information in Aspose.Cells | Prevent hyperlink export in CSV output using Aspose.Cells .NET library
// Tags: Aspose.Cells CSV export without hyperlinks | TxtSaveOptions SaveFormat.Csv C# | Excel to CSV conversion ignoring hyperlink data | C# workbook to CSV using Aspose.Cells | exclude external links in CSV output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions with SaveFormat.Csv, notes that the ExportHyperlinks property is unavailable, and saves the workbook as a CSV file while handling any runtime exceptions.
class WorkbookToCsvConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure CSV save options (use TxtSaveOptions for CSV format)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            // Note: ExportHyperlinks property is not available in this version of Aspose.Cells.
            // If needed, use CsvSaveOptions (inherits TxtSaveOptions) which may expose additional settings.

            // Save the workbook as CSV using the configured options
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook successfully converted to CSV: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
