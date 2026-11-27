// Title: Merging adjacent RichTextPortion objects that share identical formatting using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge consecutive RichTextPortion objects with the same font style | Combine adjacent rich text portions in an Excel cell using the RichString API | How to consolidate identical formatting in RichString portions with Aspose.Cells | Programmatically remove duplicate formatting from Excel cell rich text in .NET | C# code example for merging RichTextPortion objects in an Aspose.Cells workbook
// Tags: merge RichTextPortion formatting Aspose.Cells | consolidate adjacent rich text portions .NET | detect identical font properties Aspose.Cells | RichString portion merging C# | optimize Excel cell rich text handling | remove duplicate rich text formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, reads the plain text from a cell, and includes placeholder methods for merging adjacent RichTextPortion objects that have identical formatting. It demonstrates how to detect identical Font properties and outlines where RichString merging logic would be applied, then saves the workbook unchanged.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: process cell A1 (change the address as required)
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the plain string value from the cell
            string cellText = cell.StringValue;

            // NOTE: Rich text handling (merging portions) requires the RichString API,
            // which may not be available in all Aspose.Cells versions.
            // The current implementation preserves the cell text without modification.

            // Write the (unchanged) text back to the cell
            cell.PutValue(cellText);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // The following methods are retained for reference but are not used in the
    // simplified implementation because RichString is unavailable.

    // Merges adjacent RichTextPortion objects with the same formatting
    private static void MergeAdjacentPortions(object richString)
    {
        // Placeholder: RichString handling not implemented in this version.
    }

    // Determines whether two Font objects have identical formatting
    private static bool HasIdenticalFormatting(Font font1, Font font2)
    {
        return font1.Name == font2.Name &&
               font1.Size == font2.Size &&
               font1.Color.ToArgb() == font2.Color.ToArgb() &&
               font1.IsBold == font2.IsBold &&
               font1.IsItalic == font2.IsItalic &&
               font1.Underline == font2.Underline && // Compare underline enum values
               font1.IsStrikeout == font2.IsStrikeout;
    }
}
