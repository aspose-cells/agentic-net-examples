// Title: Load an XLSX workbook with Aspose.Cells LoadOptions and save it as PDF using C#
// Common Searches: how to load an Excel file with LoadOptions and convert it to PDF using Aspose.Cells in C# | C# example for converting XLSX to PDF with Aspose.Cells while handling missing input file | Aspose.Cells SaveFormat.Pdf usage in .NET to export a workbook to PDF | convert Excel workbook to PDF without chart rendering using Aspose.Cells C# | exception handling when saving a workbook as PDF with Aspose.Cells
// Tags: Aspose.Cells LoadOptions Xlsx to PDF conversion | C# workbook export to PDF Aspose.Cells | Excel to PDF conversion with SaveFormat.Pdf | handling missing Excel file Aspose.Cells C# | exception safe PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an input XLSX file, loads it with Aspose.Cells LoadOptions, and then saves the workbook directly as a PDF using SaveFormat.Pdf, with basic exception handling for load or save errors.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputFile = "input.xlsx";
        string outputFile = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // LoadOptions configured to load the workbook (default loads all content)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Save the workbook as PDF
            workbook.Save(outputFile, SaveFormat.Pdf);

            Console.WriteLine($"PDF file successfully created at '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during loading or saving
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
