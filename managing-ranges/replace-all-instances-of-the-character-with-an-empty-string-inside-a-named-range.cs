// Title: How to remove asterisk characters from string cells within a named range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace asterisk in cells of a named range | remove * characters from Excel named range using Aspose.Cells .NET | iterate over named range cells and clean string values with Aspose.Cells | Aspose.Cells delete specific characters from workbook cells C# | clean up special characters in Excel named range with Aspose.Cells
// Tags: asterisk removal from named range Aspose.Cells | Excel string cell sanitization C# | named range cell iteration Aspose.Cells | character stripping from workbook cells .NET | named range value transformation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, obtains the named range "MyNamedRange", iterates through each cell in that range, and for any string cell replaces all asterisk (*) characters with an empty string before saving the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (use GetRangeByName)
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyNamedRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyNamedRange' not found.");
                return;
            }

            // Iterate through all cells in the named range
            foreach (Cell cell in namedRange)
            {
                // Process only string values
                if (cell.Value != null && cell.Value is string)
                {
                    string original = (string)cell.Value;
                    string replaced = original.Replace("*", string.Empty);
                    cell.PutValue(replaced);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
