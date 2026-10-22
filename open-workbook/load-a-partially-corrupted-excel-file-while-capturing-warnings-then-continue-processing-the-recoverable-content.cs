// Title: Recover a partially corrupted XLSX workbook in C# using Aspose.Cells LoadOptions and capture load warnings
// Common Searches: C# Aspose.Cells load corrupted XLSX file and retrieve warning messages | How to recover a partially damaged Excel workbook with Aspose.Cells LoadOptions in .NET | Capture load warnings while opening a corrupted workbook using Aspose.Cells in C# | Save a recovered workbook after loading a corrupted XLSX with Aspose.Cells | Retrieve recoverable content from a damaged Excel file using Aspose.Cells .NET API
// Tags: Aspose.Cells LoadOptions corrupted XLSX recovery | capture load warnings Aspose.Cells .NET | recoverable content from damaged Excel workbook | save recovered workbook Aspose.Cells C# | handle partially corrupted Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a partially corrupted XLSX file, configures LoadOptions for the XLSX format, loads the workbook with Aspose.Cells while capturing any load warnings, accesses the first worksheet, writes a value to cell A1, and saves the recovered workbook to a new file, handling exceptions gracefully.
class Program
{
    static void Main()
    {
        // Path to the partially corrupted Excel file
        string inputPath = "corrupted.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Configure load options for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet and read its name
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"First worksheet name: {sheet.Name}");

            // Example modification: write a value to cell A1
            sheet.Cells["A1"].PutValue("Processed after recovery");

            // Save the recovered workbook to a new file
            string outputPath = "recovered_output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Recovered workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
