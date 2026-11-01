// Title: Create a union range from A1:B2 and D4:E5 and apply bold font style using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create union range from addresses A1:B2 and D4:E5 | How to apply a bold font to a union range using WorksheetCollection.CreateUnionRange | Example of formatting non‑contiguous cells with StyleFlag in Aspose.Cells | Saving a workbook after styling a union range in Aspose.Cells .NET | C# try‑catch pattern for workbook generation with Aspose.Cells
// Tags: WorksheetCollection.CreateUnionRange with StyleFlag | bold font style for union range Aspose.Cells | format non‑contiguous cells C# Aspose.Cells | save workbook after applying style Aspose.Cells | exception handling in Aspose.Cells workbook creation

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates a new workbook, builds a union range covering cells A1:B2 and D4:E5 on the first worksheet, defines a bold font style with a StyleFlag, applies the style to the entire union range, ensures the output directory exists, and saves the workbook to output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the worksheets collection
            WorksheetCollection worksheets = workbook.Worksheets;

            // Create a union range from the specified addresses on the first worksheet (index 0)
            UnionRange unionRange = worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

            // Create a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Define a style flag to apply only the font bold attribute
            StyleFlag flag = new StyleFlag
            {
                FontBold = true
            };

            // Apply the bold style to all cells in the union range
            unionRange.ApplyStyle(boldStyle, flag);

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
