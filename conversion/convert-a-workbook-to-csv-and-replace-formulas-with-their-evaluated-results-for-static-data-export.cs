// Title: Convert an Excel workbook to CSV with formulas evaluated using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export worksheet to CSV after calculating formulas | How to save an .xlsx as static CSV using TxtSaveOptions in Aspose.Cells | Aspose.Cells calculate all formulas before saving as CSV in .NET | Convert .xlsx to .csv with evaluated formulas using Aspose.Cells C# | Save active worksheet as UTF-8 CSV with Aspose.Cells in C#
// Tags: Aspose.Cells calculate formulas before CSV export | TxtSaveOptions CSV configuration UTF-8 C# | Workbook.Save as CSV with evaluated values | Excel to CSV static data export Aspose.Cells | C# convert .xlsx to .csv using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving; // For TxtSaveOptions

// The example loads an Excel workbook, forces formula calculation so cells contain static values, configures TxtSaveOptions for UTF-8 CSV output with a comma separator, and saves the active worksheet as a CSV file, handling missing files and other exceptions.
class WorkbookToCsvExporter
{
    static void Main()
    {
        // Paths for input workbook and output CSV
        string sourcePath = "input.xlsx";
        string csvPath = "output.csv";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Evaluate all formulas so that cells contain static values
            workbook.CalculateFormula();

            // Configure CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8, // Use UTF-8 encoding
                Separator = ','           // Column separator
                // ConvertNumericDataToText property is not available in this version
            };

            // Save the active worksheet as CSV with evaluated values
            workbook.Save(csvPath, csvOptions);
            Console.WriteLine($"Workbook successfully exported to CSV: {csvPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during export: {ex.Message}");
        }
    }
}
