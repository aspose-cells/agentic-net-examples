// Title: Convert an Excel workbook to PDF and define the Subject built-in document property with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set Subject property before saving workbook as PDF | How to add built-in document properties to an Excel file when converting to PDF using Aspose.Cells | Aspose.Cells example for XLSX to PDF conversion with custom metadata in .NET | Set Excel workbook Subject metadata for PDF output with Aspose.Cells C# | Saving an Excel workbook as PDF while preserving document properties using Aspose.Cells
// Tags: Aspose.Cells workbook subject metadata | Excel to PDF conversion with metadata C# | built-in document properties Aspose.Cells | save workbook as PDF with custom subject | C# convert XLSX to PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file with Aspose.Cells, assigns a value to the BuiltInDocumentProperties.Subject for categorization, and then saves the workbook as a PDF, handling errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Set the Subject property for categorization purposes
            workbook.BuiltInDocumentProperties.Subject = "Financial Report Q4";

            // Convert and save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
