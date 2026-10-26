// Title: Load an Excel workbook, keep only the defined name "Total", and export it to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to retain only a specific defined name when converting Excel to PDF | C# remove all named ranges except 'Total' before saving workbook as PDF using Aspose.Cells | load Excel file with Aspose.Cells and export to PDF without including charts | filter defined names in Aspose.Cells workbook programmatically C# | convert .xlsx to .pdf in C# while preserving only selected named ranges
// Tags: keep selected defined name Aspose.Cells C# | export workbook to PDF Aspose.Cells | remove unwanted named ranges Aspose.Cells | load Excel without charts Aspose.Cells | save Excel as PDF C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, removes every defined name that is not exactly "Total", and then saves the modified workbook as a PDF, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (charts are loaded; disabling them can be added if supported by the library version)
            Workbook workbook = new Workbook(inputPath);

            // Keep only defined names that are exactly "Total"
            var definedNames = workbook.Worksheets.Names;
            for (int i = definedNames.Count - 1; i >= 0; i--)
            {
                if (!string.Equals(definedNames[i].Text, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    definedNames.RemoveAt(i);
                }
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
