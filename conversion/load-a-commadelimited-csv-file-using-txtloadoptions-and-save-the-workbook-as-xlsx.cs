// Title: Convert a comma‑delimited CSV file to XLSX using Aspose.Cells TxtLoadOptions in C#
// Common Searches: c# load csv with comma separator using Aspose.Cells TxtLoadOptions | aspnet convert csv to xlsx with Aspose.Cells custom delimiter | how to specify separator for CSV import in Aspose.Cells C# | save workbook as xlsx after loading csv with Aspose.Cells | aspose.cells txtloadoptions example for csv to xlsx conversion c#
// Tags: Aspose.Cells TxtLoadOptions CSV import | C# CSV to XLSX conversion Aspose.Cells | custom CSV delimiter Aspose.Cells | save workbook as Xlsx Aspose.Cells | load CSV with comma separator Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the CSV file, configures TxtLoadOptions with a comma separator, loads the CSV into a Workbook, and saves it as an XLSX file, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Path to the destination XLSX file
            string xlsxPath = "output.xlsx";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Input file not found: {csvPath}");
                return;
            }

            // Configure load options for a comma‑delimited CSV file
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.CSV)
            {
                Separator = ',' // Ensure comma is used as the delimiter
            };

            // Load the CSV file into a workbook using the specified options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Save the workbook in XLSX format
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion successful: '{csvPath}' → '{xlsxPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
