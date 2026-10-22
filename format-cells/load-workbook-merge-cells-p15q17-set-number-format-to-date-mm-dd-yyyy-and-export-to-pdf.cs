// Title: Merge cells P15:Q17, apply a mm-dd-yyyy date format, and save the workbook as PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge a range and set a custom date format before exporting to PDF | C# code to apply mm-dd-yyyy number format to merged cells with Aspose.Cells | Export an Excel worksheet with merged cells to PDF while preserving date formatting using Aspose.Cells | How to merge cells P15 to Q17 and format them as dates in Aspose.Cells for .NET | Saving a workbook that contains merged date cells as PDF with Aspose.Cells C#
// Tags: merge cells Aspose.Cells C# | custom date number format Aspose.Cells | export workbook to PDF Aspose.Cells | range styling merged cells Aspose.Cells | mm-dd-yyyy format Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file with Aspose.Cells, merges the range P15:Q17 on the first worksheet, creates a style that uses the custom date format "mm-dd-yyyy" and applies it to the merged range, then saves the workbook as a PDF file, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells P15:Q17 (zero‑based indices: row 14, column 15, 3 rows, 2 columns)
            sheet.Cells.Merge(14, 15, 3, 2);

            // Create a style with the desired date format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm-dd-yyyy";

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("P15:Q17");
            mergedRange.ApplyStyle(dateStyle, new StyleFlag { NumberFormat = true });

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
