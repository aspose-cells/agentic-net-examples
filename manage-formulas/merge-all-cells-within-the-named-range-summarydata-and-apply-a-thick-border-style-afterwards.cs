// Title: How to merge cells in a named range and apply a thick border using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a workbook, retrieves a named range with GetRangeByName, merges the range, and sets a thick border on all sides using Aspose.Cells. | Show how to create a custom Style with thick borders and apply it to a merged named range in an Excel file with Aspose.Cells for .NET. | Adapt the example to accept input file path, named range name, and output file path as parameters while merging and styling the range.
// Common Searches: Aspose.Cells C# merge cells of a named range and add thick borders | GetRangeByName example with border styling in Aspose.Cells | Apply custom border style to merged range using Aspose.Cells for .NET | How to merge a named range and set border thickness in C# Excel library
// Tags: merge named range Aspose.Cells C# | thick border style Aspose.Cells | GetRangeByName API Aspose.Cells | style merged cells Excel .NET | custom cell border Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, retrieves the named range "SummaryData" with GetRangeByName, merges all cells in that range, applies a thick border on every side using a custom style, and saves the result to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "SummaryData"
            // Use GetRangeByName (the correct Aspose.Cells API) instead of GetNamedRange
            Aspose.Cells.Range summaryRange = workbook.Worksheets.GetRangeByName("SummaryData");
            if (summaryRange == null)
            {
                Console.WriteLine("Named range 'SummaryData' was not found.");
                return;
            }

            // Merge all cells within the named range
            summaryRange.Merge();

            // Create a style with thick borders on all sides
            Style thickBorderStyle = workbook.CreateStyle();
            thickBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            thickBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

            // Apply the style to the merged range
            StyleFlag flag = new StyleFlag { All = true };
            summaryRange.ApplyStyle(thickBorderStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
