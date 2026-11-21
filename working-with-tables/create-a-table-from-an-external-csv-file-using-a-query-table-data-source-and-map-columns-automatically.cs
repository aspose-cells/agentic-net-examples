// Title: Load an external CSV into an Aspose.Cells workbook, auto‑fit columns, rename the sheet, and save as XLSX using C#
// Common Searches: how to load a CSV file into Aspose.Cells workbook in C# | Aspose.Cells C# example for converting CSV to XLSX with auto column width | C# code to rename worksheet after importing CSV with Aspose.Cells | using LoadOptions to read CSV and save as Excel file in Aspose.Cells | Aspose.Cells auto‑fit columns after loading CSV data in C#
// Tags: load csv with Aspose.Cells LoadOptions C# | convert csv to xlsx Aspose.Cells | auto fit columns Aspose.Cells worksheet | rename worksheet after csv import Aspose.Cells | save workbook as xlsx Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for a CSV file, loads it into an Aspose.Cells Workbook using LoadOptions for CSV format, renames the first worksheet to "CSVData", auto‑fits all columns, and then saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the external CSV file (adjust as needed)
            string csvPath = @"C:\Data\sample.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            // Load the CSV file into a workbook using LoadOptions for CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "CSVData";

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the workbook to an XLSX file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully as {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
