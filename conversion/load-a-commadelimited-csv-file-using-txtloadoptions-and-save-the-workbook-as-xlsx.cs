// Title: Load a comma‑delimited CSV file with TxtLoadOptions and convert it to XLSX using Aspose.Cells for .NET (C#)
// Common Searches: C# load CSV file with custom separator using Aspose.Cells TxtLoadOptions | Aspose.Cells convert comma separated values to Excel workbook in .NET | How to save a workbook as XLSX after loading a CSV with Aspose.Cells | Set separator for CSV import in Aspose.Cells C# example | Convert input.csv to output.xlsx using Aspose.Cells library
// Tags: TxtLoadOptions with comma delimiter | Aspose.Cells CSV import settings | Workbook export to Xlsx format | C# Aspose.Cells load options usage | CSV to Excel conversion using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the CSV file exists, configures TxtLoadOptions to use a comma separator, loads the CSV into an Aspose.Cells Workbook, and then saves the workbook as an XLSX file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source CSV file
            string csvFilePath = "input.csv";

            // Verify the CSV file exists
            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine($"CSV file not found: {csvFilePath}");
                return;
            }

            // Configure load options for a comma‑delimited CSV file
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.CSV)
            {
                Separator = ',' // Ensure comma is used as the delimiter
            };

            // Load the CSV into a workbook
            Workbook workbook = new Workbook(csvFilePath, loadOptions);

            // Save the workbook as an XLSX file
            string xlsxFilePath = "output.xlsx";
            workbook.Save(xlsxFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {xlsxFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
