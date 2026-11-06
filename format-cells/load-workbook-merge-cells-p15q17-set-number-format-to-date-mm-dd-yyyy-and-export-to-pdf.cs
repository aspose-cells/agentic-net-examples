// Title: Merge cells P15:Q17, set a mm-dd-yyyy date format, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net how to merge cells P15 Q17 and apply a custom mm-dd-yyyy date format with Aspose.Cells before saving as PDF | apply date style to a merged range in Excel using Aspose.Cells .NET | export an Excel worksheet that contains merged cells to PDF with Aspose.Cells in C# | set custom number format on merged cells and generate PDF using Aspose.Cells library
// Tags: merge cell range Aspose.Cells C# | apply mm-dd-yyyy date style Aspose.Cells | save workbook as PDF Aspose.Cells | style merged range Aspose.Cells | formatted cells PDF export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using CellsRange = Aspose.Cells.Range;

// The example loads an existing workbook, merges the range P15:Q17 on the first worksheet, creates a style with the custom date format "mm-dd-yyyy" and applies it to the merged cells, then saves the workbook as a PDF file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells P15:Q17 (rows and columns are zero‑based)
            sheet.Cells.Merge(14, 15, 3, 2); // P15 is row 14, column 15

            // Create a style with the desired date format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm-dd-yyyy";

            // Apply the style to the merged range using a Range object
            CellsRange mergedRange = sheet.Cells.CreateRange("P15", "Q17");
            mergedRange.SetStyle(dateStyle);

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
