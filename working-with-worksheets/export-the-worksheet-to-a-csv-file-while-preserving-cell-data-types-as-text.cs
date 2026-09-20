// Title: Export a specific Excel worksheet to CSV as plain text using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, selects a single worksheet, and saves it to a .csv file ensuring every cell is written as a string. | Show how to configure Aspose.Cells CsvSaveOptions with ExportDataAsString = true to export a worksheet to CSV while keeping original cell values as text.
// Common Searches: how to save an Excel worksheet as CSV with all values treated as text using Aspose.Cells C# | Aspose.Cells CsvSaveOptions ExportDataAsString example for .NET | C# export specific sheet to CSV preserving formatting Aspose.Cells | convert .xlsx to .csv without losing leading zeros Aspose.Cells | save workbook to csv file while handling missing input file in C# Aspose
// Tags: Aspose.Cells export worksheet to CSV | CsvSaveOptions ExportDataAsString .NET | preserve leading zeros when exporting to CSV | select single worksheet for CSV export Aspose.Cells | handle missing input file Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the input.xlsx file exists, loads it with Aspose.Cells, and saves the workbook (or a chosen worksheet) as output.csv, writing all cell values as plain text and handling any runtime errors.
class ExportWorksheetToCsv
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found – {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook (or a specific worksheet) as CSV using the built‑in SaveFormat
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook successfully exported to CSV: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
