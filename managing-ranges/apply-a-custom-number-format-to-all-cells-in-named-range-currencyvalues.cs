// Title: Apply a custom currency number format to every cell in the named range "CurrencyValues" with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to set a custom currency format for a named range | apply number format to all cells in a named range using Aspose.Cells .NET | Aspose.Cells example applying StyleFlag to a named range in an Excel workbook | set custom number format "$#,##0.00" for range retrieved by GetRangeByName in C# | format cells as currency in Excel via Aspose.Cells named range
// Tags: custom currency number format Aspose.Cells | apply style to named range C# | StyleFlag NumberFormat usage Aspose.Cells | GetRangeByName Excel workbook Aspose.Cells | range.ApplyStyle custom style C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, obtains the named range "CurrencyValues" with GetRangeByName, creates a Style with the custom currency format "$#,##0.00", enables only the NumberFormat flag, applies the style to every cell in the range via ApplyStyle, and saves the updated workbook.
class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range associated with the named range "CurrencyValues"
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("CurrencyValues");
            if (range == null)
            {
                Console.WriteLine("Named range 'CurrencyValues' not found.");
                return;
            }

            // Create a style with the desired custom number format (e.g., currency)
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "$#,##0.00";

            // Define which style attributes to apply (only number format in this case)
            StyleFlag flag = new StyleFlag
            {
                NumberFormat = true
            };

            // Apply the custom style to all cells in the named range
            range.ApplyStyle(customStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
