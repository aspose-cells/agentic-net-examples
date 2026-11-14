// Title: How to convert an XLSX workbook to PDF using Aspose.Cells in C# (.NET)
// Common Searches: C# Aspose.Cells convert Excel .xlsx file to PDF programmatically | Aspose.Cells save workbook as PDF with SaveFormat.Pdf in .NET | Check if Excel file exists before converting to PDF using Aspose.Cells C# | Handle exceptions when exporting Excel to PDF with Aspose.Cells | Example code for loading XLSX and exporting to PDF with Aspose.Cells
// Tags: Aspose.Cells XLSX to PDF conversion C# | SaveFormat.Pdf workbook export Aspose.Cells | C# file existence validation before Aspose.Cells load | Aspose.Cells exception handling pattern | Excel to PDF rendering using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample program checks that the input.xlsx file is present, loads it into an Aspose.Cells Workbook, and then saves the workbook as output.pdf using the SaveFormat.Pdf option, with proper exception handling for robust conversion.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the XLSX workbook from file
            Workbook workbook = new Workbook(inputFile);

            // Save the workbook as a PDF document
            workbook.Save(outputFile, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputFile}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
