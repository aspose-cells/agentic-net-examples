// Title: Export every worksheet from an Excel workbook into a single CSV file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, sets the TxtSaveOptions.ExportAllSheets flag, and writes all worksheets to one CSV file. | Demonstrate how to configure Aspose.Cells TxtSaveOptions for CSV to combine multiple sheets into a single output stream in a .NET application.
// Common Searches: Aspose.Cells how to save all Excel sheets as one CSV in C# | C# combine multiple worksheets into a single CSV using Aspose.Cells | TxtSaveOptions ExportAllSheets true example for CSV export | Save workbook with all sheets to one CSV file Aspose .NET | Export Excel workbook to combined CSV file programmatically
// Tags: Aspose.Cells CSV export all sheets | TxtSaveOptions CSV multiple worksheets | combine Excel worksheets into single CSV .NET | save workbook as combined CSV Aspose | C# Aspose.Cells CSV export options

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input.xlsx file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions with SaveFormat.Csv and ExportAllSheets enabled, and saves all worksheets together into output.csv while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure CSV save options to export all worksheets into one file
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.ExportAllSheets = true;

            // Save the combined CSV
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully saved as CSV to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
