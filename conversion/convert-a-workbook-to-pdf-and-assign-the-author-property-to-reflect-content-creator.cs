// Title: Convert an Excel workbook to PDF and set the author property using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook author before exporting to PDF | convert xlsx to pdf with custom author metadata using Aspose.Cells | C# example for setting built-in document properties and saving as PDF with Aspose.Cells | how to add author information to PDF generated from Excel with Aspose.Cells .NET | Aspose.Cells save workbook as PDF and include document properties in the output
// Tags: Aspose.Cells document property author assignment | Excel workbook PDF export with metadata | C# Aspose.Cells SaveFormat.Pdf example | Built-in document properties manipulation Aspose.Cells | PDF generation from XLSX using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file, assigns a value to the workbook's BuiltInDocumentProperties.Author, and then saves the workbook as a PDF using Aspose.Cells' SaveFormat.Pdf, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the author property of the workbook
            workbook.BuiltInDocumentProperties.Author = "John Doe";

            // Convert and save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
