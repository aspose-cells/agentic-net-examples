// Title: How to apply a custom number format that adds a plus sign to positive percentages in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# custom number format for positive percentages with leading plus sign | apply +0.00%;-0.00%;0.00% format to a range using Aspose.Cells .NET | format Excel percentages with a plus sign for positive values in C# | set custom percentage format for positive, negative and zero values in Aspose.Cells | range.ApplyStyle example with custom number format in Aspose.Cells
// Tags: custom percentage format with plus sign Aspose.Cells | apply style to cell range Aspose.Cells .NET | range.ApplyStyle custom number format C# | Excel workbook save with custom number format | percentage formatting leading plus sign C#

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, inserts numeric values, defines the custom format '+0.00%;-0.00%;0.00%' to show a leading plus sign for positive percentages, applies the style to cells A1:A3 via range.ApplyStyle, and saves the file as FormattedPercentages.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Insert sample numeric values (these will be shown as percentages)
            sheet.Cells["A1"].PutValue(0.1234);   // 12.34%
            sheet.Cells["A2"].PutValue(-0.0567); // -5.67%
            sheet.Cells["A3"].PutValue(0);       // 0%

            // Define a custom number format:
            //   +0.00%  -> positive values with a leading plus sign
            //   -0.00%  -> negative values with a minus sign
            //   0.00%   -> zero values without sign
            string customFormat = "+0.00%;-0.00%;0.00%";

            // Create a style that uses the custom format
            Style style = workbook.CreateStyle();
            style.Custom = customFormat;

            // Indicate that the number format should be applied
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true; // Apply number formatting

            // Apply the style to the range A1:A3 (use fully qualified Aspose.Cells.Range)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A3");
            range.ApplyStyle(style, flag);

            // Define output file path
            string outputPath = "FormattedPercentages.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
