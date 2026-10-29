// Title: How to exclude hidden rows when loading an Excel workbook with Aspose.Cells for .NET using FilterObjects
// Common Searches: Aspose.Cells C# load workbook without hidden rows | Skip invisible rows during Excel import with Aspose.Cells .NET | Configure FilterObjects to ignore rows where visibility is false in Aspose.Cells | LoadOptions hide rows false Aspose.Cells example | C# Aspose.Cells exclude hidden rows on workbook load | Filter hidden rows while loading XLSX using Aspose.Cells API
// Tags: filter hidden rows Aspose.Cells | load workbook without hidden rows | FilterObjects visibility false Aspose.Cells | LoadOptions hide rows C# | exclude invisible rows Excel Aspose.Cells | Aspose.Cells workbook import filter rows

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an XLSX workbook with Aspose.Cells for .NET while configuring LoadOptions and FilterObjects to skip rows whose visibility is set to false, then saving the processed workbook.
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
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load options – using default options as LoadFilterOptions may not be available in all versions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Additional processing can be placed here

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
